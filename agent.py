# agent.py — Mac Agent controlling Microsoft PowerPoint via AppleScript
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
import subprocess, json, hmac, hashlib, os, re, time, pathlib

app = FastAPI()
SECRET = os.getenv("DISPATCH_SECRET", "dev-secret")
SAVE = os.path.expanduser("~/Documents/JarvisDecks")
os.makedirs(SAVE, exist_ok=True)


def verify(sig: str, payload: dict):
    mac = hmac.new(
        SECRET.encode(),
        json.dumps(payload, sort_keys=True).encode(),
        hashlib.sha256,
    ).hexdigest()
    if not hmac.compare_digest(mac, sig or ""):
        raise HTTPException(status_code=401, detail="bad signature")


def _as_escape(s: str) -> str:
    """Escape for AppleScript string literals."""
    if s is None:
        return ""
    # normalize smart quotes → straight, escape backslashes / quotes
    s = s.replace("’", "'").replace("“", '"').replace("”", '"')
    s = s.replace("\\", "\\\\").replace('"', '\\"')
    return s


def make_deck(title: str, bullets: list[str]):
    ts = time.strftime("%Y%m%d-%H%M%S")

    # sanitize file name from title (no backslashes in f-string expr)
    safe_title = re.sub(r"[^A-Za-z0-9_\- ]", "_", title)
    safe_title = safe_title.strip().replace(" ", "_")

    base = pathlib.Path(SAVE) / f"{safe_title}-{ts}"
    pptx = str(base.with_suffix(".pptx"))
    pdf = str(base.with_suffix(".pdf"))

    esc_title = _as_escape(title)
    # AppleScript uses \r to separate lines in a text frame (bullets)
    esc_bullets = "\\r".join([_as_escape(b) for b in bullets])

    script = f'''
    tell application "Microsoft PowerPoint"
      activate
      set p to make new presentation
      set s1 to make new slide at end of slides of p with properties {{layout:layoutTitle}}
      tell s1
        set the text range of (text frame of shape 1) to "{esc_title}"
      end tell
      set s2 to make new slide at end of slides of p with properties {{layout:layoutText}}
      tell s2
        set the text range of (text frame of shape 2) to "{esc_bullets}"
      end tell
      save as p filename "{pptx}"
      save as p filename "{pdf}" file format save as PDF
    end tell
    '''
    proc = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "Unknown AppleScript error").strip()
        raise HTTPException(status_code=500, detail=f"AppleScript error: {err}")
    return pptx, pdf


@app.post("/command")
async def command(req: Request):
    payload = await req.json()
    verify(req.headers.get("X-Signature", ""), payload)
    cmd = (payload.get("command") or "").strip()
    low = cmd.lower()

    if not low.startswith("ppt create"):
        return JSONResponse({"message": "Unsupported command"}, status_code=400)

    # Accept title/bullets in single or double quotes
    # Examples:
    #   ppt create title:"Quick Deck" bullets:"A;B;C"
    #   ppt create title:'Quick Deck' bullets:'A;B;C'
    title_m = re.search(r"title\s*:\s*(['\"])(.*?)\1", cmd, flags=re.I | re.S)
    bullets_m = re.search(r"bullets\s*:\s*(['\"])(.*?)\1", cmd, flags=re.I | re.S)

    title = title_m.group(2) if title_m else "Auto Deck"
    raw_bullets = bullets_m.group(2) if bullets_m else "Overview; Highlights; Risks; Next steps"
    bullets = [b.strip() for b in raw_bullets.split(";") if b.strip()]

    pptx, pdf = make_deck(title, bullets)
    return {"message": f"Deck ready\nPPTX: {pptx}\nPDF: {pdf}"}


@app.get("/health")
def health():
    return {"ok": True, "save_dir": SAVE, "time": time.strftime("%Y-%m-%d %H:%M:%S")}
