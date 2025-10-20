import os
import httpx

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PPLX_API_KEY = os.getenv("PPLX_API_KEY")


def _guard() -> str:
    return (
        "Return ONLY a unified diff (git apply format). "
        "No prose, no fences. Under 400 lines. Touch only allowed paths. "
        "If unsure, return empty."
    )


def call_openai(prompt: str, model="gpt-4o-mini") -> str:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY missing")
    r = httpx.post(
        "https://api.openai.com/v1/chat/completions",
        headers={"Authorization": f"Bearer {OPENAI_API_KEY}"},
        json={
            "model": model,
            "messages": [
                {
                    "role": "system",
                    "content": "You generate STRICT unified diffs only.",
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.2,
            "max_tokens": 1600,
        },
        timeout=60,
    )
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def call_perplexity(prompt: str, model="sonar-small") -> str:
    if not PPLX_API_KEY:
        raise RuntimeError("PPLX_API_KEY missing")
    r = httpx.post(
        "https://api.perplexity.ai/chat/completions",
        headers={
            "Authorization": f"Bearer {PPLX_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": [
                {
                    "role": "system",
                    "content": "You generate STRICT unified diffs only.",
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.0,
            "max_tokens": 1200,
        },
        timeout=60,
    )
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def ask_for_patch(instructions: str, diff_hint: str = "", provider="openai") -> str:
    prompt = f"{_guard()}\n\nContext:\n{instructions}\n\nOptional hints:\n{diff_hint}"
    return call_perplexity(prompt) if provider == "perplexity" else call_openai(prompt)
