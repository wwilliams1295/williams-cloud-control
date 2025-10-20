# plugins/send_pdf.py
from typing import Any

try:
    from mailer import send_email
except Exception as e:
    # Defer error until run() so the module can still be discovered
    send_email = None
    _import_error = e
else:
    _import_error = None

name = "send_pdf"
description = "Email PDF(s)/files via SMTP"
permissions: list[str] = ["smtp:send", "fs:read:*"]


def run(
    to: str,
    files: list,
    subject: str = "Requested document",
    body: str = "Please see attached.",
) -> dict[str, Any]:
    if _import_error:
        return {"ok": False, "error": f"mailer import failed: {_import_error}"}
    if not files:
        return {"ok": False, "error": "No files provided"}
    send_email(to=to, subject=subject, body=body, files=files)
    return {"ok": True, "sent_to": to, "count": len(files)}
