# plugins/edgar_pull.py — fetch latest 10-Q/10-K HTML from SEC
import os
import pathlib
import requests
import time
from typing import Any
from core.capabilities import REGISTRY, CapabilitySpec

UA = os.getenv("SEC_USER_AGENT", "ExampleContact you@example.com")
DATA_DIR = pathlib.Path("./data")
DATA_DIR.mkdir(exist_ok=True)


def _cik_from_ticker(ticker: str) -> str:
    url = "https://data.sec.gov/api/xbrl/taxonomy/2024/us-gaap.json"  # lightweight ping to set UA
    requests.get(url, headers={"User-Agent": UA}, timeout=10)
    requests.get(
        "https://data.sec.gov/api/xbrl/companyfacts/CIK0000320193.json",
        headers={"User-Agent": UA},
        timeout=10,
    )
    # ^ priming call (SEC rate limit friendliness). Real mapping below:
    tmap = requests.get(
        "https://www.sec.gov/files/company_tickers.json",
        headers={"User-Agent": UA},
        timeout=20,
    ).json()
    for _, rec in tmap.items():
        if rec.get("ticker", "").upper() == ticker.upper():
            return f"{int(rec['cik_str']):010d}"
    raise ValueError(f"Ticker not found: {ticker}")


def run(ticker: str, prefer: str = "10-Q") -> dict[str, Any]:
    """
    Returns {'ok': bool, 'html_path': str, 'form': '10-Q'|'10-K', 'filing_date': 'YYYY-MM-DD'}
    """
    cik = _cik_from_ticker(ticker)
    sub = requests.get(
        f"https://data.sec.gov/submissions/CIK{cik}.json",
        headers={"User-Agent": UA},
        timeout=20,
    ).json()
    # find latest preferred form, else fallback to 10-K
    forms = sub.get("filings", {}).get("recent", {})
    candidates = []
    for i, form in enumerate(forms.get("form", [])):
        if form in (prefer, "10-K"):
            candidates.append(
                (form, forms["accessionNumber"][i], forms["filingDate"][i])
            )
    if not candidates:
        return {"ok": False, "error": "No suitable filing found."}
    form, acc, fdate = candidates[0]
    acc_nodashes = acc.replace("-", "")
    idx_url = (
        f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc_nodashes}/index.json"
    )
    idx = requests.get(idx_url, headers={"User-Agent": UA}, timeout=20).json()
    # choose the main HTML doc
    doc = next(
        (d for d in idx["directory"]["item"] if d["name"].endswith(".htm")), None
    )
    if not doc:
        return {"ok": False, "error": "No HTML document found in filing."}
    raw_url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc_nodashes}/{doc['name']}"
    html = requests.get(raw_url, headers={"User-Agent": UA}, timeout=30).text
    out_path = DATA_DIR / f"{ticker.upper()}_{form}_{fdate}.html"
    out_path.write_text(html, encoding="utf-8")
    time.sleep(0.2)  # SEC friendly
    return {"ok": True, "html_path": str(out_path), "form": form, "filing_date": fdate}


REGISTRY.register(
    CapabilitySpec(
        name="edgar_pull",
        version="1.0.0",
        description="Fetch latest SEC filing HTML for a ticker (prefers form 10-Q else 10-K).",
        entrypoint=run,
        inputs_schema={
            "type": "object",
            "required": ["ticker"],
            "properties": {"ticker": {"type": "string"}, "prefer": {"type": "string"}},
        },
        outputs_schema={"type": "object"},
        example={"tool": "edgar_pull", "args": {"ticker": "NVDA", "prefer": "10-Q"}},
    )
)
