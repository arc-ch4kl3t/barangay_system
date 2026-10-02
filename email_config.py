import os
import json
import urllib.request
import urllib.error
from dotenv import load_dotenv

load_dotenv()

BREVO_API_KEY = os.getenv("BREVO_API_KEY", "").strip()
SENDER_EMAIL = os.getenv("GMAIL_ADDRESS", "lorainenina40@gmail.com").strip()
SENDER_NAME = "Barangay Information System"

if BREVO_API_KEY:
    print("[EMAIL CONFIG SUCCESS] BREVO_API_KEY is configured.")
else:
    print("[EMAIL CONFIG ERROR] BREVO_API_KEY environment variable is MISSING or EMPTY!")

class SafeConfigDict(dict):
    def __getitem__(self, key):
        if key in ("sender_email", "email", "user", "username"):
            return SENDER_EMAIL
        if key in ("sender_name", "name"):
            return SENDER_NAME
        return super().get(key, "")

GMAIL_CONFIG = SafeConfigDict({
    "sender_email": SENDER_EMAIL,
    "email": SENDER_EMAIL,
    "password": "",
    "sender_password": "",
    "sender_name": SENDER_NAME
})

def send_gmail_message(msg):
    """
    Sends email via Brevo HTTP API (Port 443) to bypass Render SMTP restrictions.
    """
    if not BREVO_API_KEY:
        print("[EMAIL ERROR] Missing BREVO_API_KEY in environment variables.")
        return False, "Email service not configured. Contact administrator."

    to_email = msg["To"]
    subject = msg["Subject"]

    html_content = ""
    text_content = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/html":
                html_content = part.get_payload(decode=True).decode("utf-8", errors="ignore")
            elif content_type == "text/plain":
                text_content = part.get_payload(decode=True).decode("utf-8", errors="ignore")
    else:
        html_content = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

    payload = {
        "sender": {"name": SENDER_NAME, "email": SENDER_EMAIL},
        "to": [{"email": to_email}],
        "subject": subject,
        "htmlContent": html_content or text_content,
    }

    headers = {
        "accept": "application/json",
        "api-key": BREVO_API_KEY,
        "content-type": "application/json"
    }

    req = urllib.request.Request(
        "https://api.brevo.com/v3/smtp/email",
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            if response.status in (200, 201):
                print(f"[EMAIL SUCCESS] Brevo API delivered reset email to {to_email}")
                return True, "Email sent successfully"
            else:
                return False, f"Brevo API error status: {response.status}"
    except urllib.error.HTTPError as exc:
        err_body = exc.read().decode('utf-8', errors='ignore')
        print(f"[EMAIL ERROR] Brevo HTTP Error {exc.code}: {err_body}")
        return False, f"Failed to send email via API: {exc.code}"
    except Exception as exc:
        print(f"[EMAIL ERROR] Network error sending via Brevo: {exc}")
        return False, f"Failed to send email: {exc}"
