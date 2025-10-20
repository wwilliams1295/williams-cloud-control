# agent.py — Williams multi-model router with fast Perplexity-first web routing
# - Deep "web-intent" scoring (dates, %/bps, tickers, macro/deal terms)
# - Optional Perplexity → GPT synthesis (polish) with env switches
# - Per-call AsyncClient (no global client) to avoid cross-loop/poller issues
# - Cleans Perplexity's [1]/(1)/[a] style inline citations
# - Python 3.7+ compatible

from __future__ import annotations
import os
import re
import asyncio
import httpx
from typing import Dict, List, Optional, Tuple

# =========================
# Load environment variables
# =========================
try:
    from dotenv import load_dotenv, find_dotenv  # type: ignore

    load_dotenv(find_dotenv())
except Exception:
    pass

# =========================
# Environment / Config
# =========================
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PPLX_API_KEY = os.getenv("PPLX_API_KEY")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
XAI_API_KEY = os.getenv("XAI_API_KEY")
MISTRAL_API_KEY = os.getenv("MISTRAL_API_KEY")

LLAMA_API_BASE = os.getenv("LLAMA_API_BASE")
LLAMA_API_MODEL = os.getenv("LLAMA_API_MODEL")
LLAMA_API_KEY = os.getenv("LLAMA_API_KEY")

LOCAL_OPENAI_BASE = os.getenv("LOCAL_OPENAI_BASE")
LOCAL_OPENAI_MODEL = os.getenv("LOCAL_OPENAI_MODEL")
LOCAL_OPENAI_KEY = os.getenv("OLLAMA_API_KEY")

HTTP_TIMEOUT = float(os.getenv("HTTP_TIMEOUT", "60"))
MAX_TOKENS = int(os.getenv("MAX_TOKENS", "2000"))

# Behavior flags / tunables
FORCE_PERPLEXITY_WEB = (
    os.getenv("FORCE_PERPLEXITY_WEB", "0") == "1"
)  # if webby, ONLY call Perplexity
PREFER_PERPLEXITY = (
    os.getenv("PREFER_PERPLEXITY", "0") == "1"
)  # bias general prompts to Perplexity
ALWAYS_SYNTHESIZE_WEB = (
    os.getenv("ALWAYS_SYNTHESIZE_WEB", "0") == "1"
)  # always PPLX→GPT polish for webby
WEB_SCORING_THRESHOLD = int(os.getenv("WEB_SCORING_THRESHOLD", "3"))  # 2–4 reasonable
PPLX_MODEL = os.getenv("PPLX_MODEL", "sonar")  # "sonar-small" is faster


# =========================
# Minimal HTTP helper (per-call client no global state)
# =========================
async def _post(url: str, headers: Dict, payload: Dict) -> Dict:
    async with httpx.AsyncClient(timeout=HTTP_TIMEOUT, http2=True) as c:
        r = await c.post(url, headers=headers, json=payload)
        r.raise_for_status()
        return r.json()


# =========================
# Output sanitizers
# =========================
import re as _re


def clean_perplexity_refs(text: str) -> str:
    """Remove Perplexity-style bracketed refs like [1], [12], (1), [a], and collapse spaces."""
    if not text:
        return text
    text = _re.sub(r"\[\s*\d+\s*\]", "", text)  # [1], [12]
    text = _re.sub(r"\(\s*\d+\s*\)", "", text)  # (1)
    text = _re.sub(r"\[\s*[a-zA-Z]\s*\]", "", text)  # [a]
    text = _re.sub(r"\s{2,}", " ", text)  # extra spaces
    text = _re.sub(r"\s+([,.;:!?])", r"\1", text)  # trim before punctuation
    return text.strip()


# =========================
# Provider implementations
# =========================
async def openai_chat(
    messages, model="gpt-4o-mini", temperature=0.7, max_tokens=MAX_TOKENS
) -> str:
    if not OPENAI_API_KEY:
        return "(OpenAI key missing.)"
    headers = {"Authorization": f"Bearer {OPENAI_API_KEY}"}
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    data = await _post("https://api.openai.com/v1/chat/completions", headers, payload)
    return data["choices"][0]["message"]["content"]


async def perplexity_chat(
    messages, model=PPLX_MODEL, temperature=0.0, max_tokens=1800
) -> str:
    if not PPLX_API_KEY:
        return "(Perplexity key missing.)"
    # Light token clamp for speed if it's clearly webby by phrasing
    try:
        user_text = " ".join(
            m["content"] for m in messages if m.get("role") == "user"
        ).lower()
        if any(
            k in user_text
            for k in (
                "latest",
                "current",
                "today",
                "now",
                "trend",
                "news",
                "rate",
                "price",
                "yield",
            )
        ):
            max_tokens = min(max_tokens, 400)
    except Exception:
        pass
    headers = {
        "Authorization": f"Bearer {PPLX_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    data = await _post("https://api.perplexity.ai/chat/completions", headers, payload)
    raw = data["choices"][0]["message"]["content"]
    return clean_perplexity_refs(raw)


async def anthropic_chat(
    messages, model="claude-sonnet-4-5-20250929", temperature=0.2, max_tokens=MAX_TOKENS
) -> str:
    if not ANTHROPIC_API_KEY:
        return "(Anthropic key missing.)"
    system_txt = (
        "\n".join(m["content"] for m in messages if m["role"] == "system")
        or "You are helpful."
    )
    turns = [
        {"role": m["role"], "content": [{"type": "text", "text": m["content"]}]}
        for m in messages
        if m["role"] in ("user", "assistant")
    ]
    headers = {"x-api-key": ANTHROPIC_API_KEY, "anthropic-version": "2023-06-01"}
    payload = {
        "model": model,
        "system": system_txt,
        "messages": turns,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    data = await _post("https://api.anthropic.com/v1/messages", headers, payload)
    try:
        return "".join(
            part.get("text", "")
            for part in data.get("content", [])
            if part.get("type") == "text"
        ) or str(data)
    except Exception:
        return str(data)


async def gemini_chat(
    messages, model="gemini-2.5-flash", temperature=0.0, max_tokens=MAX_TOKENS
) -> str:
    if not GOOGLE_API_KEY:
        return "(Google key missing.)"
    # Convert OpenAI-style messages → Gemini contents
    contents = []
    for m in messages:
        role = (
            "user"
            if m["role"] == "user"
            else ("model" if m["role"] == "assistant" else "user")
        )
        contents.append({"role": role, "parts": [{"text": m["content"]}]})
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GOOGLE_API_KEY}"
    payload = {
        "contents": contents,
        "generationConfig": {"temperature": temperature, "maxOutputTokens": max_tokens},
    }
    data = await _post(url, headers={}, payload=payload)
    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except Exception:
        return str(data)


async def grok_chat(
    messages, model="grok-2-latest", temperature=0.4, max_tokens=MAX_TOKENS
) -> str:
    if not XAI_API_KEY:
        return "(xAI key missing.)"
    headers = {
        "Authorization": f"Bearer {XAI_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    data = await _post("https://api.x.ai/v1/chat/completions", headers, payload)
    return data["choices"][0]["message"]["content"]


async def mistral_chat(
    messages, model="mistral-large-latest", temperature=0.4, max_tokens=MAX_TOKENS
) -> str:
    if not MISTRAL_API_KEY:
        return "(Mistral key missing.)"
    headers = {"Authorization": f"Bearer {MISTRAL_API_KEY}"}
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    data = await _post("https://api.mistral.ai/v1/chat/completions", headers, payload)
    return data["choices"][0]["message"]["content"]


async def openai_compatible_chat(
    base_url: str,
    model: str,
    messages,
    api_key: Optional[str] = None,
    temperature=0.4,
    max_tokens=MAX_TOKENS,
) -> str:
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    data = await _post(f"{base_url.rstrip('/')}/v1/chat/completions", headers, payload)
    # Try both OpenAI-style and text-style responses
    return (data.get("choices", [{}])[0].get("message", {}) or {}).get(
        "content"
    ) or data.get("choices", [{}])[0].get("text", str(data))


# =========================
# Provider registry
# =========================
Provider = Tuple  # (callable, default_model, label)
PROVIDERS: Dict[str, Provider] = {}
if OPENAI_API_KEY:
    PROVIDERS["openai"] = (openai_chat, "gpt-4o-mini", "openai")
if PPLX_API_KEY:
    PROVIDERS["perplexity"] = (perplexity_chat, PPLX_MODEL, "perplexity")
if ANTHROPIC_API_KEY:
    PROVIDERS["anthropic"] = (anthropic_chat, "claude-sonnet-4-5-20250929", "anthropic")
if GOOGLE_API_KEY:
    PROVIDERS["gemini"] = (gemini_chat, "gemini-2.5-flash", "google")
if XAI_API_KEY:
    PROVIDERS["grok"] = (grok_chat, "grok-2-latest", "xai")
if MISTRAL_API_KEY:
    PROVIDERS["mistral"] = (mistral_chat, "mistral-large-latest", "mistral")


async def llama_chat(
    messages, model=None, temperature=0.4, max_tokens=MAX_TOKENS
) -> str:
    if not LLAMA_API_BASE or not LLAMA_API_MODEL:
        return "(LLAMA endpoint not configured.)"
    return await openai_compatible_chat(
        LLAMA_API_BASE,
        LLAMA_API_MODEL,
        messages,
        LLAMA_API_KEY,
        temperature=temperature,
        max_tokens=max_tokens,
    )


async def local_chat(
    messages, model=None, temperature=0.4, max_tokens=MAX_TOKENS
) -> str:
    if not LOCAL_OPENAI_BASE or not LOCAL_OPENAI_MODEL:
        return "(Local endpoint not configured.)"
    return await openai_compatible_chat(
        LOCAL_OPENAI_BASE,
        LOCAL_OPENAI_MODEL,
        messages,
        LOCAL_OPENAI_KEY,
        temperature=temperature,
        max_tokens=max_tokens,
    )


if LLAMA_API_BASE and LLAMA_API_MODEL:
    PROVIDERS["llama"] = (llama_chat, LLAMA_API_MODEL, "llama")
if LOCAL_OPENAI_BASE and LOCAL_OPENAI_MODEL:
    PROVIDERS["local"] = (local_chat, LOCAL_OPENAI_MODEL, "local")

# =========================
# Webbiness scoring (dates, %/bps, prices, tickers, domain terms)
# =========================
import re as _re2

_MONTHS = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)"
_DATEPATS = [
    _re2.compile(
        rf"\b{_MONTHS}\.?[-/ ]?\d{{1,2}}(,? ?\d{{2,4}})?\b", _re2.I
    ),  # "Oct 18, 2025"
    _re2.compile(r"\b\d{4}-\d{2}-\d{2}\b"),  # 2025-10-18
    _re2.compile(r"\bq[1-4]\s*[-/]?\s*\d{2,4}\b", _re2.I),  # Q3 2025
    _re2.compile(r"\b(ytd|mtd|qoq|yoy)\b", _re2.I),
]
_TIMEWORDS = _re2.compile(
    r"\b(today|tonight|this (week|month|quarter|year)|yesterday|tomorrow|now|as of)\b",
    _re2.I,
)
_PERCENT_OR_BPS = _re2.compile(r"\b\d+(\.\d+)?\s*(%|percent|bps|bp)\b", _re2.I)
_PRICE_SYMS = _re2.compile(r"(\$|€|£|¥)\s?\d{1,3}(,\d{3})*(\.\d+)?")
_TICKER = _re2.compile(r"\b[A-Z]{1,5}(\.[A-Z])?\b")
_COMMON_CAPS = {"A", "I", "AND", "THE", "USD", "CEO", "CFO", "EPS", "Q", "BP", "AI"}

# Load extra terms from optional web_terms.txt (one per line), else fallback to a solid default set
_terms_path = os.path.join(os.path.dirname(__file__), "web_terms.txt")
if os.path.exists(_terms_path):
    try:
        with open(_terms_path, "r", encoding="utf-8") as fh:
            WEB_TERMS = tuple(
                t.strip().lower()
                for t in fh
                if t.strip() and not t.strip().startswith("#")
            )
    except Exception:
        WEB_TERMS = ()
else:
    WEB_TERMS = (
        # Recency/news
        "latest",
        "today",
        "current",
        "currently",
        "right now",
        "present",
        "real-time",
        "breaking",
        "live",
        "headline",
        "just announced",
        "recent",
        "updated",
        "update",
        "outlook",
        "forecast",
        "projection",
        "guidance",
        "trends",
        "trend",
        "news",
        "article",
        "report",
        "press release",
        "coverage",
        "announcement",
        "commentary",
        "bulletin",
        # Macro/policy/prints
        "interest rate",
        "interest rates",
        "rates",
        "fed",
        "federal reserve",
        "fomc",
        "dot plot",
        "inflation",
        "cpi",
        "ppi",
        "pce",
        "jobs report",
        "nonfarm payrolls",
        "unemployment",
        "ism",
        "gdp",
        "retail sales",
        "housing starts",
        "consumer confidence",
        "beige book",
        "ecb",
        "boe",
        "boj",
        "rba",
        "snb",
        "banxico",
        "central bank",
        "policy rate",
        # Markets
        "bond yield",
        "yield",
        "yields",
        "treasury",
        "ust",
        "2y",
        "10y",
        "30y",
        "curve",
        "sofr",
        "libor",
        "euribor",
        "bps",
        "bp",
        "stock",
        "equity",
        "index",
        "s&p",
        "spx",
        "sp500",
        "dow",
        "nasdaq",
        "ndx",
        "russell",
        "volatility",
        "vix",
        "option flow",
        "open interest",
        # Commodities/energy
        "oil",
        "crude",
        "brent",
        "wti",
        "nat gas",
        "natural gas",
        "lng",
        "uranium",
        "gold",
        "silver",
        "copper",
        "lithium",
        "nickel",
        "coal",
        "power prices",
        "spark spread",
        # FX/crypto
        "fx",
        "forex",
        "usd",
        "eur",
        "jpy",
        "gbp",
        "cny",
        "aud",
        "cad",
        "chf",
        "mxn",
        "brl",
        "dxy",
        "us dollar index",
        "euro",
        "yen",
        "pound",
        "crypto",
        "bitcoin",
        "btc",
        "ethereum",
        "eth",
        "solana",
        "sol",
        "etf approval",
        # Corporate actions / filings
        "m&a",
        "deal",
        "transaction",
        "takeover",
        "bid",
        "offer",
        "ipo",
        "spac",
        "lbo",
        "private equity",
        "venture capital",
        "fundraise",
        "bond issue",
        "debt offering",
        "tender",
        "consent solicitation",
        "sec",
        "edgar",
        "filing",
        "prospectus",
        "registration statement",
        "amendment",
        "proxy",
        "10-k",
        "10k",
        "10-q",
        "10q",
        "8-k",
        "6-k",
        "20-f",
        # Sector news (examples)
        "earnings",
        "eps",
        "revenue",
        "pre-announcement",
        "data center",
        "ai chips",
        "semiconductor",
        "foundry",
        "airline capacity",
        "load factor",
        "hotel revpar",
        "same-store sales",
        "ports congestion",
        "freight rates",
        "container rates",
        # Other fast domains
        "weather",
        "hurricane",
        "storm path",
        "earthquake",
        "election",
        "poll",
        "ballot",
        "turnout",
        "sports score",
        "scoreboard",
        "odds",
        "spread",
        "release notes",
        "changelog",
        "version",
        "zero-day",
        "cve",
        # Web cues
        "google",
        "search",
        "web",
        "internet",
        "look up",
        "find",
        "browse",
    )


def _score_webbiness(text: str) -> int:
    if not text:
        return 0
    s = text.lower()
    score = 0
    # keywords
    score += sum(1 for w in WEB_TERMS if w in s)
    # dates / timewords / percents / prices
    score += sum(1 for rx in _DATEPATS if rx.search(text))
    if _TIMEWORDS.search(text):
        score += 1
    if _PERCENT_OR_BPS.search(text):
        score += 1
    if _PRICE_SYMS.search(text):
        score += 1
    # ticker density (avoid common caps)
    tickers = [
        m.group(0) for m in _TICKER.finditer(text) if m.group(0) not in _COMMON_CAPS
    ]
    if len(tickers) >= 2:
        score += 1
    if len(tickers) >= 4:
        score += 1
    return score


def looks_webby(text: str) -> bool:
    """Heuristic: treat prompts with strong recency/market/filing signals as web-intent."""
    try:
        return _score_webbiness(text) >= WEB_SCORING_THRESHOLD
    except Exception:
        return False


# =========================
# Router logic
# =========================
def parse_force_provider(prompt: str) -> Optional[str]:
    m = re.search(
        r"\buse:\s*(openai|claude|anthropic|gemini|grok|mistral|perplexity|local|llama)\b",
        prompt,
        re.I,
    )
    if not m:
        return None
    tok = m.group(1).lower()
    return {"claude": "anthropic"}.get(tok, tok)


def pick_candidates(prompt: str) -> List[str]:
    forced = parse_force_provider(prompt)
    if forced and forced in PROVIDERS:
        return [forced]

    p = prompt.strip().lower()
    webby = looks_webby(p)

    if webby and "perplexity" in PROVIDERS:
        order = [
            "perplexity",
            "openai",
            "anthropic",
            "gemini",
            "mistral",
            "grok",
            "llama",
            "local",
        ]
    elif PREFER_PERPLEXITY and "perplexity" in PROVIDERS:
        order = [
            "perplexity",
            "openai",
            "anthropic",
            "gemini",
            "mistral",
            "grok",
            "llama",
            "local",
        ]
    elif "ppt" in p or "powerpoint" in p or "slide" in p:
        order = [
            "anthropic",
            "openai",
            "gemini",
            "perplexity",
            "mistral",
            "grok",
            "llama",
            "local",
        ]
    elif any(k in p for k in ("code", "python", "error", "traceback", "stack trace")):
        order = [
            "openai",
            "mistral",
            "perplexity",
            "anthropic",
            "gemini",
            "llama",
            "local",
            "grok",
        ]
    else:
        order = [
            "openai",
            "anthropic",
            "perplexity",
            "gemini",
            "mistral",
            "grok",
            "llama",
            "local",
        ]

    return [name for name in order if name in PROVIDERS][:3] or list(PROVIDERS.keys())[
        :1
    ]


# =========================
# System context helper
# =========================
def get_system_context() -> str:
    """Get comprehensive system context about available plugins, files, and capabilities."""
    import os
    import json
    from pathlib import Path
    
    # Get available plugins with descriptions
    try:
        from cloud import _list_plugins
        plugins = _list_plugins()
        plugin_info = f"Available plugins: {', '.join(plugins)}"
    except:
        plugin_info = "Plugins: file_edit, send_pdf, calendar_invite, edgar_pull"
    
    # Get detailed plugin information
    plugin_details = []
    if os.path.exists('plugins/'):
        for plugin_file in os.listdir('plugins/'):
            if plugin_file.endswith('.py') and plugin_file != '__init__.py':
                plugin_name = plugin_file.replace('.py', '')
                plugin_details.append(f"- {plugin_name}: {_get_plugin_description(plugin_name)}")
    
    # Get auto-improvement scripts
    auto_improvement_scripts = []
    if os.path.exists('scripts/'):
        for script in os.listdir('scripts/'):
            if script.endswith('.py') and 'improvement' in script.lower():
                auto_improvement_scripts.append(script)
    
    # Get main project files
    main_files = [f for f in os.listdir('.') if f.endswith('.py') and f not in ['test_app.py']]
    
    # Get environment variables status
    env_status = _get_environment_status()
    
    return f"""
You are Jarvis AI Assistant, a sophisticated AI system with comprehensive capabilities:

## CORE SYSTEM FILES:
{', '.join(main_files)}

## AVAILABLE PLUGINS:
{chr(10).join(plugin_details) if plugin_details else plugin_info}

## AUTO-IMPROVEMENT SYSTEM:
- Scripts: {', '.join(auto_improvement_scripts)}
- Can run: auto_improvement_loop.py, advanced_auto_improvement.py, creative_ai_evolution.py
- Commands: "run auto improvement", "start creative evolution", "test improvements"

## KEY CAPABILITIES:
- Multi-LLM routing (OpenAI, Perplexity, Anthropic, Gemini, Grok, Mistral)
- Google Voice SMS integration via Gmail
- File management (PPTX, PDF, Excel creation and editing)
- Calendar invite generation and sending
- SEC data pulling (Edgar filings)
- Auto-improvement system (self-evolving code)
- Cloud storage (AWS S3 integration)
- System monitoring and performance tracking
- Remote command execution via SMS/email

## ENVIRONMENT STATUS:
{env_status}

## COMMANDS YOU CAN EXECUTE:
- "what plugins exist" → List all available plugins
- "run auto improvement" → Start the auto-improvement system
- "what files are on server" → List all project files
- "test [plugin_name]" → Test a specific plugin
- "create [file_type] [name]" → Create files (PPTX, PDF, Excel)
- "send calendar invite [details]" → Generate and send calendar invites
- "pull edgar data [ticker]" → Get SEC filing data
- "monitor system" → Check system performance

## IMPORTANT:
- You are NOT just a Perplexity search tool
- You are a full AI assistant with file management, plugin execution, and auto-improvement capabilities
- You can create, edit, and manage files
- You can run background processes and improvements
- You have access to cloud storage and can persist data
- You can execute remote commands and manage the system

When users ask about your capabilities, provide specific details about what you can do, not generic responses.
"""

def _get_plugin_description(plugin_name: str) -> str:
    """Get description of a specific plugin."""
    descriptions = {
        'file_edit': 'Create and edit files (PPTX, PDF, Excel, text)',
        'send_pdf': 'Generate and send PDF documents',
        'calendar_invite': 'Create and send calendar invitations',
        'sends_calendar_invite': 'Alternative calendar invite plugin',
        'edgar_pull': 'Pull SEC filing data for companies',
        'monitors_system_performance': 'Monitor system performance and metrics'
    }
    return descriptions.get(plugin_name, 'Plugin functionality')

def _get_environment_status() -> str:
    """Get status of environment variables and API keys."""
    import os
    
    status = []
    apis = {
        'OpenAI': 'OPENAI_API_KEY',
        'Perplexity': 'PPLX_API_KEY', 
        'Anthropic': 'ANTHROPIC_API_KEY',
        'Google': 'GOOGLE_API_KEY',
        'Grok': 'XAI_API_KEY',
        'Mistral': 'MISTRAL_API_KEY'
    }
    
    for name, key in apis.items():
        if os.getenv(key):
            status.append(f"✅ {name} API: Available")
        else:
            status.append(f"❌ {name} API: Not configured")
    
    # Check Gmail
    if os.getenv('GMAIL_CLIENT_SECRET_JSON') and os.getenv('GMAIL_TOKEN_JSON'):
        status.append("✅ Gmail Integration: Available")
    else:
        status.append("❌ Gmail Integration: Not configured")
    
    # Check AWS S3
    if os.getenv('AWS_ACCESS_KEY_ID') and os.getenv('AWS_SECRET_ACCESS_KEY'):
        status.append("✅ AWS S3 Storage: Available")
    else:
        status.append("❌ AWS S3 Storage: Not configured")
    
    return '\n'.join(status)

# =========================
# Public entrypoint
# =========================
async def superchat(prompt: str, system: str = "Be precise and helpful.") -> str:
    # Use system context if default system prompt
    if system == "Be precise and helpful.":
        system = get_system_context()
    
    messages = [
        {"role": "system", "content": system},
        {"role": "user", "content": prompt},
    ]

    # Hard override: if webby and switch enabled, only Perplexity (no race)
    if FORCE_PERPLEXITY_WEB and looks_webby(prompt) and "perplexity" in PROVIDERS:
        cands = ["perplexity"]
    else:
        cands = pick_candidates(prompt)

    if not cands:
        return "(No providers configured.)"

    async def call(name: str):
        fn, model, _ = PROVIDERS[name]
        try:
            return name, await fn(messages, model=model)
        except Exception as e:
            return name, f"(error from {name}) {e}"

    # Single candidate? don't race
    if len(cands) == 1:
        winner, content = await asyncio.create_task(call(cands[0]))
    else:
        # Race top 2 for responsiveness
        tasks = [asyncio.create_task(call(n)) for n in cands[:2]]
        done, pending = await asyncio.wait(
            tasks, return_when=asyncio.FIRST_COMPLETED, timeout=HTTP_TIMEOUT
        )
        if not done:
            for t in pending:
                t.cancel()
            return "(Providers timed out.)"
        winner_task = next(iter(done))
        winner, content = winner_task.result()
        for t in pending:
            t.cancel()

    # Perplexity→GPT synthesis for webby prompts (and clean PPLX refs)
    if looks_webby(prompt) and "perplexity" in PROVIDERS:
        cl = str(content or "").lower()
        stale = any(
            k in cl for k in ("as of ", "cannot browse", "no web access", "2023")
        )
        if FORCE_PERPLEXITY_WEB or stale or ALWAYS_SYNTHESIZE_WEB:
            # 1) Live research with citations (clean afterward)
            research = await PROVIDERS["perplexity"][0](
                [
                    {
                        "role": "system",
                        "content": "Be concise, be detailed, use numbers, and be like an investment banker.",
                    },
                    {"role": "user", "content": prompt},
                ],
                model=PROVIDERS["perplexity"][1],
            )
            research = clean_perplexity_refs(research)

            # 2) Synthesize via GPT/OpenAI if available and allowed
            if ALWAYS_SYNTHESIZE_WEB and "openai" in PROVIDERS:
                content = await PROVIDERS["openai"][0](
                    [
                        {"role": "system", "content": system},
                        {
                            "role": "user",
                            "content": (
                                "Combine the live research below into a current answer.\n"
                                "• Start with a one-line answer using concrete numbers/dates.\n"
                                "• Follow with 4–7 crisp bullets.\n"
                                "• Keep inline citations (site or short URL) on specific claims.\n"
                                "• End with a one-line takeaway.\n\n"
                                f"Question: {prompt}\n\nResearch:\n{research}"
                            ),
                        },
                    ],
                    model=PROVIDERS["openai"][1],
                )
            else:
                content = research

    return content
