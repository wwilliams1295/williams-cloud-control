import os
import smtplib
import mimetypes
import pathlib
from email.message import EmailMessage

SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")
SMTP_TLS = os.getenv("SMTP_TLS", "1") == "1"
FROM_EMAIL = os.getenv("FROM_EMAIL", SMTP_USER)
MAX_TOTAL_BYTES = int(os.getenv("MAIL_MAX_TOTAL_BYTES", str(20 * 1024 * 1024)))


class AttachmentTooLarge(Exception):
    pass


class MissingCredentials(Exception):
    pass


def _attach_files(msg: EmailMessage, files):
    total = 0
    for f in files or []:
        p = pathlib.Path(f)
        if not p.exists():
            raise FileNotFoundError(f"Attachment not found: {p}")
        data = p.read_bytes()
        total += len(data)
        if total > MAX_TOTAL_BYTES:
            raise AttachmentTooLarge(f"Attachments exceed {MAX_TOTAL_BYTES} bytes")
        ctype, _ = mimetypes.guess_type(p.name)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=p.name)


def send_email(to, subject, body, files=None, body_html=None, cc=None, bcc=None):
    if not (SMTP_USER and SMTP_PASS and FROM_EMAIL):
        raise MissingCredentials("SMTP_USER/SMTP_PASS/FROM_EMAIL not set")
    if isinstance(to, str):
        to = [to]
    cc = cc or []
    bcc = bcc or []
    msg = EmailMessage()
    msg["From"] = FROM_EMAIL
    msg["To"] = ", ".join(to)
    msg["Subject"] = subject
    if cc:
        msg["Cc"] = ", ".join(cc)
    if body_html:
        msg.set_content(body or "See HTML part.")
        msg.add_alternative(body_html, subtype="html")
    else:
        msg.set_content(body or "")
    _attach_files(msg, files)
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
        if SMTP_TLS:
            s.starttls()
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg, to_addrs=(to + cc + bcc))
