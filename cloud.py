# cloud.py — Williams GPT-5 Local Node with Codes + Excel + PPT + AI + Web + Email I/O (cloud + Mac-agent fallback)
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import PlainTextResponse
from typing import Optional, List, Tuple, Dict
import os, json, hmac, hashlib, requests, re, random, string, traceback, time, mimetypes
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from bs4 import BeautifulSoup
from duckduckgo_search import DDGS
from openai import OpenAI

# Excel
from pathlib import Path
from openpyxl import Workbook

# Email (SMTP replies)
import smtplib
from email.message import EmailMessage
from email.utils import parseaddr

# Optional cloud PPTX generator
from pptx import Presentation
from pptx.util import Inches, Pt

load_dotenv()
app = FastAPI()

# ============== CONFIG ==============
SECRET         = os.getenv("DISPATCH_SECRET", "dev-secret")
MAC_AGENT_URL  = os.getenv("MAC_AGENT_URL", "http://127.0.0.1:8787")
WIN_AGENT_URL  = os.getenv("WIN_AGENT_URL")  # optional
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TZ             = os.getenv("APP_TIMEZONE", "America/Chicago")
DB_PATH        = os.getenv("CODES_DB", os.path.expanduser("~/jarvis-demo/commands.json"))
SHEETS_DIR     = os.getenv("SHEETS_DIR", os.path.expanduser("~/Documents/JarvisSheets"))
Path(SHEETS_DIR).mkdir(parents=True, exist_ok=True)

# Email config
EMAIL_FROM      = os.getenv("EMAIL_FROM", "")
EMAIL_WHITELIST = [e.strip().lower() for e in (os.getenv("EMAIL_WHITELIST","").split(",")) if e.strip()]
SMTP_HOST       = os.getenv("SMTP_HOST", "")
SMTP_PORT       = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER       = os.getenv("SMTP_USER", "")
SMTP_PASS       = os.getenv("SMTP_PASS", "")
SMTP_DEBUG      = os.getenv("SMTP_DEBUG","0") == "1"

# Optional footer (e.g., for promos/newsletters)
COMPANY_ADDRESS = os.getenv("COMPANY_ADDRESS", "")

# Greeting behavior (avoid repeating the “authenticated” line)
GREET_COOLDOWN_SECONDS = int(os.getenv("GREET_COOLDOWN_SECONDS", "300"))  # 5 min default
_last_greet_at: Dict[str, float] = {}  # keyed by sender_id

# Cloud output dir for generated PPTX, etc.
CLOUD_OUT_DIR = Path(os.getenv("CLOUD_OUT_DIR", os.path.expanduser("~/Documents/JarvisCloud")))
CLOUD_OUT_DIR.mkdir(parents=True, exist_ok=True)

client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

ALLOWED_NUMBERS = {
    "+15613891295": {"name": "Chris", "role": "admin"},
}

# ============== UTIL ==============
def now_str():
    try:
        return datetime.now(ZoneInfo(TZ)).strftime("%A, %B %-d, %Y %I:%M %p %Z")
    except Exception:
        return datetime.now().strftime("%A, %B %d, %Y %I:%M %p")

def twiml(msg: str) -> PlainTextResponse:
    return PlainTextResponse(
        f'<?xml version="1.0"?><Response><Message>{msg}</Message></Response>',
        media_type="application/xml"
    )

def sign_payload(d: dict) -> str:
    return hmac.new(SECRET.encode(), json.dumps(d, sort_keys=True).encode(), hashlib.sha256).hexdigest()

def maybe_greeting(sender_key: str, name: str) -> str:
    """Return greeting once per GREET_COOLDOWN_SECONDS per sender."""
    now = time.time()
    last = _last_greet_at.get(sender_key, 0.0)
    if now - last < GREET_COOLDOWN_SECONDS:
        return ""  # suppress repeated greeting
    _last_greet_at[sender_key] = now
    call_signs = ["Williams Echo-Nine", "Williams Core-One", "Palm Node", "Houston Command"]
    call_sign = random.choice(call_signs)
    return (
        f"{name}, you have been authenticated to the Williams Secured Cloud Control Server. "
        f"Connection uplink established with {call_sign}. "
        f"All systems nominal. How can I assist you today?"
    )

# Email send helper (HARDENED)
def send_email(to_addr: str, subject: str, body: str, attachments: Optional[List[str]] = None) -> str:
    if not (SMTP_HOST and SMTP_PORT and SMTP_USER and SMTP_PASS and EMAIL_FROM):
        return "(Email disabled: SMTP settings missing)"
    if COMPANY_ADDRESS:
        body = f"{body}\n\n—\n{COMPANY_ADDRESS}"
    to_email = parseaddr(to_addr)[1] or to_addr
    msg = EmailMessage()
    msg["From"] = EMAIL_FROM
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)
    for path in (attachments or []):
        if not os.path.isfile(path):
            continue
        ctype, _ = mimetypes.guess_type(path)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(path))
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            if SMTP_DEBUG:
                s.set_debuglevel(1)
            s.ehlo(); s.starttls(); s.ehlo()
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg, from_addr=EMAIL_FROM, to_addrs=[to_email])
        return "Email sent."
    except Exception as e:
        return f"Email send error: {e}"

# Find file paths mentioned in reply text (to attach automatically)
def _extract_paths(text: str) -> List[str]:
    paths: List[str] = []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith(("PPTX:", "PDF:", "XLSX:")):
            p = line.split(":", 1)[1].strip()
            if os.path.isfile(p):
                paths.append(p)
    return paths

# ============== LIGHT MEMORY (prompt seasoning) ==============
from collections import deque
MEMORY = deque(maxlen=20)
def remember(user_text: str, ai_text: str):
    MEMORY.append({"user": user_text, "ai": ai_text})
def memory_text():
    if not MEMORY: return "(no prior context)"
    lines = []
    for m in MEMORY:
        lines.append(f"User: {m['user']}")
        lines.append(f"AI: {m['ai']}")
    return "\n".join(lines[-12:])

# ============== OPENAI ==============
def ask_gpt(prompt: str, name: str = "Operator") -> str:
    if not client:
        return f"(Local mode) {name}, system time is {now_str()}. OpenAI not configured."
    try:
        completion = client.chat.completions.create(
            model="gpt-4o-mini",  # upgrade to "gpt-4o" if you want
            messages=[
                {"role":"system","content":
                 "You are Jarvis, the Williams Secured Cloud Control Server AI (codename Echo-Nine). "
                 "Cinematic yet composed tone. Be precise, confident, and helpful."
                },
                {"role":"system","content": f"Current date/time: {now_str()}."},
                {"role":"system","content": f"Conversation memory (latest):\n{memory_text()}"},
                {"role":"user","content": prompt},
            ],
            temperature=0.6,
        )
        return (completion.choices[0].message.content or "").strip()
    except Exception as e:
        return f"OpenAI error: {e}"

# ============== WEB SEARCH + ARTICLE ==============
def search_web(query: str, max_results: int = 5) -> List[Dict]:
    try:
        ddgs = DDGS()
        return list(ddgs.text(query, max_results=max_results))
    except Exception as e:
        return [{"title":"Web search error","body":str(e),"href":""}]

def format_search_results(results: List[Dict]) -> str:
    if not results: return "No search results found."
    out = []
    for i, r in enumerate(results, start=1):
        out.append(f"({i}) {r.get('title','Untitled')}\n{r.get('body','')}\n🔗 {r.get('href','')}")
    return "\n\n".join(out)

def fetch_article(url: str) -> str:
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Williams Cloud Node)"}
        r = requests.get(url, headers=headers, timeout=12)
        soup = BeautifulSoup(r.text, "html.parser")
        text = " ".join(p.get_text(" ", strip=True) for p in soup.find_all("p"))[:6000]
        if not text.strip():
            return "The page content appears empty or blocked. Try another URL."
        return ask_gpt(f"Summarize clearly in 4–6 bullets:\n{text}")
    except Exception as e:
        return f"Article fetch error: {e}"

# ============== FACT LOOKUP (example) ==============
def fetch_live_fact(query: str) -> str:
    try:
        q = query.lower()
        if "president" in q and ("united states" in q or "us" in q or "u.s." in q):
            r = requests.get(
                "https://en.wikipedia.org/api/rest_v1/page/summary/President_of_the_United_States",
                timeout=6,
            )
            data = r.json()
            extract = data.get("extract","")
            sentence = extract.split(".")[0]
            return f"As of {now_str()}, {sentence}."
    except Exception as e:
        return f"Fact lookup error: {e}"
    return ""

# ============== COMMAND CODES REGISTRY ==============
def _ensure_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    if not os.path.exists(DB_PATH):
        with open(DB_PATH, "w") as f:
            json.dump({"aliases":{}, "codes":{}}, f)

def _load_db():
    _ensure_db()
    with open(DB_PATH, "r") as f:
        return json.load(f)

def _save_db(db):
    with open(DB_PATH, "w") as f:
        json.dump(db, f, indent=2)

def _rand_code(n=4) -> str:
    return "".join(random.choices(string.ascii_uppercase + string.digits, k=n))

def save_code(alias: str, body: str) -> Tuple[str,str]:
    db = _load_db()
    code = _rand_code()
    while code in db["codes"]:
        code = _rand_code()
    db["aliases"][alias] = {"body": body, "code": code, "created": now_str()}
    db["codes"][code] = alias
    _save_db(db)
    return alias, code

def list_codes() -> List[Tuple[str,str]]:
    db = _load_db()
    return [(a, meta.get("code","----")) for a, meta in db["aliases"].items()]

def resolve_code(key: str) -> Optional[str]:
    db = _load_db()
    k = key.strip()
    if k.startswith("#"): k = k[1:]
    alias = db["codes"].get(k)
    if alias:
        return db["aliases"][alias]["body"]
    if k in db["aliases"]:
        return db["aliases"][k]["body"]
    return None

def delete_code(key: str) -> bool:
    db = _load_db()
    k = key.strip()
    if k.startswith("#"): k = k[1:]
    alias = db["codes"].get(k)
    if alias:
        db["codes"].pop(k, None)
        db["aliases"].pop(alias, None)
        _save_db(db)
        return True
    if k in db["aliases"]:
        code = db["aliases"][k].get("code")
        if code: db["codes"].pop(code, None)
        db["aliases"].pop(k, None)
        _save_db(db)
        return True
    return False

# ============== EXECUTORS ==============
def create_excel_file(name: str = "Quick Sheet", cols: Optional[List[str]] = None) -> str:
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe = re.sub(r"[^A-Za-z0-9_\- ]", "_", name).strip().replace(" ", "_")
    path = Path(SHEETS_DIR) / f"{safe}-{ts}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    if cols:
        ws.append(cols)
    wb.save(path)
    return str(path)

def exec_excel_command(s: str) -> str:
    m = re.search(r"name:'([^']*)'", s, flags=re.I)
    name = (m.group(1) if m else "Quick Sheet").strip()
    m = re.search(r"cols:'([^']*)'", s, flags=re.I)
    cols = [c.strip() for c in (m.group(1).split(";") if m else []) if c.strip()]
    out = create_excel_file(name=name, cols=cols)
    return f"Excel ready\nXLSX: {out}"

# ---- Cloud PPTX generator (runs even if Mac is OFF)
def make_cloud_pptx(title: str, bullets_str: str) -> str:
    prs = Presentation()
    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title or "Auto Deck"
    # Bullet slide
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "Overview"
    tf = slide2.shapes.placeholders[1].text_frame
    tf.clear()
    bullets = [b.strip() for b in bullets_str.split(";") if b.strip()]
    for i, b in enumerate(bullets):
        if i == 0:
            tf.text = b
        else:
            tf.add_paragraph().text = b
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe = re.sub(r"[^A-Za-z0-9_\- ]", "_", title or "Auto Deck").strip().replace(" ", "_")
    path = CLOUD_OUT_DIR / f"{safe}-{ts}.pptx"
    prs.save(path)
    return str(path)

def exec_deck_command(s: str, to_win: bool = False) -> str:
    """Try Mac/Win agent first; on failure or offline, generate PPTX in cloud."""
    import json
    lower = s.strip().lower()

    # Build title & bullets string
    if lower.startswith("ppt create "):
        title_m   = re.search(r"title\s*:\s*['\"](.*?)['\"]", s, flags=re.I|re.S)
        bullets_m = re.search(r"bullets\s*:\s*['\"](.*?)['\"]", s, flags=re.I|re.S)
        title = (title_m.group(1) if title_m else "Quick Deck")
        bullets_str = (bullets_m.group(1) if bullets_m else "Overview; Key drivers; Risks; Next steps")
    else:
        text = re.sub(r"^(deck:|win:)\s*", "", s, flags=re.I).strip()
        title = "Quick Deck"
        if ";" in text:
            bullets_str = text
        else:
            ai_bullets = ask_gpt(f"Generate 4–6 concise slide bullets about: {text}")
            cleaned = ai_bullets.replace("\n"," ").replace("•"," ").replace(" - "," ").strip()
            if ";" not in cleaned:
                parts = [p.strip(" .") for p in cleaned.split(".") if p.strip()]
                cleaned = "; ".join(parts[:6]) if parts else "Overview; Key drivers; Risks; Next steps"
            bullets_str = cleaned

    # Try agent (Mac or Win)
    try:
        payload = {"command": f"ppt create title:{json.dumps(title)} bullets:{json.dumps(bullets_str)}"}
        headers = {"X-Signature": sign_payload(payload)}
        agent   = (WIN_AGENT_URL if (to_win and WIN_AGENT_URL) else MAC_AGENT_URL).rstrip("/")
        url     = f"{agent}/command"
        r = requests.post(url, json=payload, headers=headers, timeout=8)
        if r.ok:
            ctype = (r.headers.get("content-type") or "").lower()
            if "application/json" in ctype:
                return r.json().get("message", "Deck complete")
            return (r.text or "Deck complete").strip()
        # Agent returned error (like AppleScript 2741) -> fall back to cloud
        if "AppleScript error" in (r.text or ""):
            path = make_cloud_pptx(title, bullets_str)
            return f"Deck complete (cloud)\nPPTX: {path}"
    except Exception:
        # Agent offline or connection error -> cloud fallback
        path = make_cloud_pptx(title, bullets_str)
        return f"Deck complete (cloud)\nPPTX: {path}"

    # Defensive final fallback
    path = make_cloud_pptx(title, bullets_str)
    return f"Deck complete (cloud)\nPPTX: {path}"

# ============== ROUTING HELPERS ==============
def should_use_chat(body: str) -> bool:
    """Default to chat; only run non-chat on explicit prefix."""
    s = (body or "").strip().lower()
    if s.startswith(("deck:", "win:", "ppt create", "excel")):
        return False
    return True

# ============== SHARED COMMAND PROCESSOR ==============
def process_message(sender_id: str, body: str, channel: str = "sms") -> str:
    # Auth & display name
    if channel == "sms":
        if sender_id not in ALLOWED_NUMBERS:
            return "Access denied. Unauthorized signal detected."
        name = ALLOWED_NUMBERS.get(sender_id, {}).get("name", "Operator")
        sender_key = sender_id
    else:  # email
        addr = parseaddr(sender_id)[1].lower()
        if EMAIL_WHITELIST and addr not in EMAIL_WHITELIST:
            return "Access denied. This email is not authorized to use Williams Cloud Control."
        name = (addr.split("@")[0].replace(".", " ").title() or "Operator")
        sender_key = addr

    # Greeting (cooldown)
    greeting = maybe_greeting(sender_key, name)
    s = (body or "").strip()
    lower = s.lower()

    # Codes control
    if lower.startswith(("code new ", "save ")):
        try:
            payload = re.sub(r"^(code new|save)\s*", "", s, flags=re.I)
            alias, raw = payload.split(":", 1)
            alias = alias.strip()
            body_for_alias = raw.strip()
            a, code = save_code(alias, body_for_alias)
            out = f"Saved code '{a}' as #{code}.\nRun with: run {a}  or  run #{code}"
            return (greeting + "\n\n" + out) if greeting else out
        except Exception as e:
            out = f"Could not save code (format: save <alias>: <body>). Error: {e}"
            return (greeting + "\n\n" + out) if greeting else out

    if lower in ("codes", "list codes", "code list"):
        pairs = list_codes()
        if not pairs:
            out = "No codes saved yet."
            return (greeting + "\n\n" + out) if greeting else out
        lines = [f"- {alias}  (#{code})" for alias, code in pairs]
        out = "Saved codes:\n" + "\n".join(lines)
        return (greeting + "\n\n" + out) if greeting else out

    if lower.startswith(("delete code ", "forget ")):
        key = re.sub(r"^(delete code|forget)\s*", "", s, flags=re.I).strip()
        ok = delete_code(key)
        out = "Deleted." if ok else "Not found."
        return (greeting + "\n\n" + out) if greeting else out

    if lower.startswith("run "):
        key = s.split(" ",1)[1].strip()
        stored = resolve_code(key)
        if not stored:
            out = f"Code '{key}' not found."
            return (greeting + "\n\n" + out) if greeting else out
        s = stored
        lower = s.lower()

    if lower.startswith("#"):
        stored = resolve_code(s)
        if not stored:
            out = f"Code '{s}' not found."
            return (greeting + "\n\n" + out) if greeting else out
        s = stored
        lower = s.lower()

    # Chat / search / scroll defaults
    if should_use_chat(s):
        cleaned = re.sub(r"^(chat:|ask:)\s*", "", s, flags=re.I).strip()

        if cleaned.lower().startswith(("search ", "google ", "find ", "look up ")):
            query = re.sub(r"^(search|google|find|look up)\s*", "", cleaned, flags=re.I)
            results = search_web(query, max_results=5)
            formatted = format_search_results(results)
            summary = ask_gpt(f"Summarize helpfully:\n{formatted}", name=name)
            out = f"Search results for '{query}':\n\n{formatted}\n\nSummary:\n{summary}"
            return (greeting + "\n\n" + out) if greeting else out

        if cleaned.lower().startswith("scroll "):
            url = cleaned.split(" ",1)[1].strip()
            article_summary = fetch_article(url)
            remember(f"Scroll: {url}", article_summary)
            out = article_summary
            return (greeting + "\n\n" + out) if greeting else out

        live = fetch_live_fact(cleaned)
        if live:
            remember(cleaned, live)
            return (greeting + "\n\n" + live) if greeting else live

        ai_reply = ask_gpt(cleaned or "Status report.", name=name)
        remember(cleaned, ai_reply)
        return (greeting + "\n\n" + ai_reply) if greeting else ai_reply

    # Excel
    if lower.startswith("excel"):
        try:
            msg = exec_excel_command(s)
            remember(s, msg)
            return (greeting + "\n\n" + msg) if greeting else msg
        except Exception as e:
            err = f"{e}\n{traceback.format_exc(limit=1)}"
            out = f"Excel exception: {err}"
            return (greeting + "\n\n" + out) if greeting else out

    # Deck (Mac/Windows agent) — only explicit
    to_win = lower.startswith("win:")
    if lower.startswith(("deck:", "win:", "ppt create ")):
        try:
            msg = exec_deck_command(s, to_win=to_win)
            remember(s, msg)
            return (greeting + "\n\n" + msg) if greeting else msg
        except Exception as e:
            err = f"{e}\n{traceback.format_exc(limit=1)}"
            out = f"Agent exception: {err}"
            return (greeting + "\n\n" + out) if greeting else out
        # SEC / EDGAR analyzer (natural language)
    if any(x in lower for x in ["sec ", "edgar ", "10-k", "10q", "10-q", "filing "]):
        try:
            # Try to detect ticker and form type automatically
            m = re.search(r"\b([A-Z]{1,5})\b", s.upper())
            ticker = m.group(1) if m else ""
            form_type = "10-Q" if "10-Q" in s.upper() or "10Q" in s.upper() else "10-K"
            if not ticker:
                out = "Please specify a ticker symbol (e.g., AAPL or XOM)."
            else:
                out = f"📊 Fetching {form_type} for {ticker} from EDGAR...\n\n" + summarize_filing(ticker, form_type=form_type)
            remember(s, out)
            return (greeting + "\n\n" + out) if greeting else out
        except Exception as e:
            err = f"SEC lookup error: {e}\n{traceback.format_exc(limit=1)}"
            return (greeting + "\n\n" + err) if greeting else err

    # Fallback
    out = "Command not recognized."
    return (greeting + "\n\n" + out) if greeting else out

# ============== HEALTH ==============
@app.get("/health")
def health():
    return {"ok": True, "time": now_str()}

# ============== SMS WEBHOOK ==============
@app.post("/twilio/sms")
async def sms_in(req: Request):
    try:
        form   = dict(await req.form())
        sender = (form.get("From") or "").strip()
        body   = (form.get("Body") or "").strip()
    except:
        return twiml("Malformed request.")
    text_reply = process_message(sender, body, channel="sms")
    return twiml(text_reply)

# ============== EMAIL INBOUND WEBHOOK ==============
@app.post("/email/inbound")
async def email_inbound(req: Request):
    """
    Works with:
      - SendGrid Inbound Parse: form-data 'from','subject','text'
      - Mailgun Routes: similar keys
      - Custom JSON: {'from','subject','text'}
    Replies by SMTP to the sender (with attachments if generated).
    """
    ctype = (req.headers.get("content-type") or "").lower()
    sender = subject = text = ""

    if "application/json" in ctype:
        data    = await req.json()
        sender  = (data.get("from") or data.get("sender") or "").strip()
        subject = (data.get("subject") or "").strip()
        text    = (data.get("text") or data.get("body") or data.get("stripped-text") or "").strip()
    else:
        form    = dict(await req.form())
        sender  = (form.get("from") or form.get("sender") or "").strip()
        subject = (form.get("subject") or "").strip()
        text    = (form.get("text") or form.get("stripped-text") or form.get("body-plain") or "").strip()

    if not sender:
        raise HTTPException(status_code=400, detail="missing sender")

    command_text = text if text else subject
    sender_addr = parseaddr(sender)[1] or sender

    reply = process_message(sender_addr, command_text, channel="email")

    # Auto-attach any generated files announced in reply text
    attachments = _extract_paths(reply)
    sent  = send_email(sender_addr, f"[Williams Cloud] Re: {subject or 'Command'}", reply, attachments=attachments)
    return {"ok": True, "sent": sent}
# ============== SEC EDGAR FINANCIAL ANALYZER ==============
import pandas as pd

def fetch_latest_filing_url(ticker: str, form_type: str = "10-K") -> str:
    """Return the main HTML URL for the latest 10-K or 10-Q filing."""
    try:
        headers = {"User-Agent": os.getenv("SEC_USER_AGENT", "JarvisCloudBot/2.0 (contact: william.c.williams@outlook.com)")}
        ticker = ticker.upper().strip()

        # Lookup CIK
        data = requests.get("https://www.sec.gov/files/company_tickers.json", headers=headers, timeout=10).json()
        cik = None
        for v in data.values():
            if v["ticker"].upper() == ticker:
                cik = str(v["cik_str"]).zfill(10)
                break
        if not cik:
            return ""

        # Get recent submissions
        sub = requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json", headers=headers, timeout=10).json()
        rec = sub.get("filings", {}).get("recent", {})
        for form, acc in zip(rec["form"], rec["accessionNumber"]):
            if form == form_type:
                base = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-','')}/{acc}-index.html"
                return base
        return ""
    except Exception:
        return ""


def parse_filing_tables(url: str) -> pd.DataFrame:
    """Extract numerical tables (Income Statement, Balance Sheet, etc.) from an HTML filing."""
    headers = {"User-Agent": os.getenv("SEC_USER_AGENT", "JarvisCloudBot/2.0 (contact: william.c.williams@outlook.com)")}
    r = requests.get(url, headers=headers, timeout=15)
    soup = BeautifulSoup(r.text, "html.parser")
    tables = soup.find_all("table")
    dfs = []
    for t in tables:
        try:
            df = pd.read_html(str(t))[0]
            if df.shape[1] >= 2:
                dfs.append(df)
        except Exception:
            continue
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def extract_financial_highlights(df: pd.DataFrame) -> dict:
    """Search the parsed DataFrame for key metrics."""
    metrics = {
        "revenue": None,
        "operating_income": None,
        "net_income": None,
        "total_assets": None,
        "total_liabilities": None,
        "shareholders_equity": None,
    }
    if df.empty:
        return metrics

    # Normalize text
    df.columns = [str(c).lower() for c in df.columns]
    for col in df.columns:
        if "description" in col or "item" in col or "account" in col:
            key_col = col
            break
    else:
        key_col = df.columns[0]

    for i, row in df.iterrows():
        text = str(row[key_col]).lower()
        valcols = [c for c in df.columns if any(x in c for x in ["amount", "value", "usd", "dollar", "current", "year"])]
        vals = [v for c, v in row.items() if c in valcols and isinstance(v, (int, float, str))]
        val = None
        if vals:
            try:
                val = float(str(vals[-1]).replace(",", "").replace("(", "-").replace(")", ""))
            except Exception:
                pass
        if not val:
            continue
        if "revenue" in text or "sales" in text:
            metrics["revenue"] = val
        elif "operating income" in text or "operating profit" in text:
            metrics["operating_income"] = val
        elif "net income" in text and "comprehensive" not in text:
            metrics["net_income"] = val
        elif "total assets" in text:
            metrics["total_assets"] = val
        elif "total liabilit" in text:
            metrics["total_liabilities"] = val
        elif "stockholders" in text or "shareholders" in text:
            metrics["shareholders_equity"] = val
    return metrics


def summarize_filing(ticker: str, form_type: str = "10-K") -> str:
    """Fetch and summarize financials with GPT commentary."""
    try:
        url = fetch_latest_filing_url(ticker, form_type)
        if not url:
            return f"❌ Could not locate latest {form_type} for {ticker}."

        df = parse_filing_tables(url)
        metrics = extract_financial_highlights(df)

        summary_prompt = (
            f"Ticker: {ticker}\n"
            f"Latest {form_type} filing: {url}\n\n"
            f"Financial snapshot (raw):\n{json.dumps(metrics, indent=2)}\n\n"
            "Please summarize the revenue growth or decline, key balance sheet changes, "
            "and any notable trends in 5 concise bullets."
        )
        return ask_gpt(summary_prompt)
    except Exception as e:
        return f"SEC summary error: {e}"
