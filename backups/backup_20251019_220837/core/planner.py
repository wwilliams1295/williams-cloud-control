# core/planner.py — ask an LLM to emit a JSON tool plan using your capabilities
import os
import json
import httpx
from typing import Any
from core.capabilities import describe_caps

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PPLX_API_KEY = os.getenv("PPLX_API_KEY")

SYSTEM = (
    "You are a planning assistant. You have tool capabilities listed below. "
    "When a user asks for something, output ONLY a JSON array of steps, where each step is "
    '{"tool":"<name>","args":{...}} using available tools. Keep it minimal and sequential. '
    "If the user email is not given, use the TO_EMAIL env hint if present; otherwise omit sending.\n\n"
    "Tools:\n"
)


def _call_openai(prompt: str) -> str:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY missing")
    r = httpx.post(
        "https://api.openai.com/v1/chat/completions",
        headers={"Authorization": f"Bearer {OPENAI_API_KEY}"},
        json={
            "model": "gpt-4o-mini",
            "messages": [
                {
                    "role": "system",
                    "content": SYSTEM + json.dumps(describe_caps(), indent=2),
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.1,
            "max_tokens": 800,
        },
        timeout=60,
    )
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def _call_perplexity(prompt: str) -> str:
    if not PPLX_API_KEY:
        raise RuntimeError("PPLX_API_KEY missing")
    r = httpx.post(
        "https://api.perplexity.ai/chat/completions",
        headers={
            "Authorization": f"Bearer {PPLX_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": "sonar-small",
            "messages": [
                {
                    "role": "system",
                    "content": SYSTEM + json.dumps(describe_caps(), indent=2),
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.0,
            "max_tokens": 700,
        },
        timeout=60,
    )
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def plan_tools(user_prompt: str, provider: str = "openai") -> list[dict[str, Any]]:
    raw = (
        _call_perplexity(user_prompt)
        if provider == "perplexity"
        else _call_openai(user_prompt)
    )
    try:
        plan = json.loads(raw)
        assert isinstance(plan, list)
        return plan
    except Exception:
        # Defensive recovery: try to extract first JSON array
        import re

        m = re.search(r"\[(?:.|\n)*\]", raw)
        if not m:
            return []
        try:
            plan = json.loads(m.group(0))
            return plan if isinstance(plan, list) else []
        except Exception:
            return []
