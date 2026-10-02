"""
Email Configuration for Password Reset
Setup instructions:
1. Go to https://myaccount.google.com/apppasswords
2. Select "Mail" and "Windows Computer" (or your device)
3. Copy the 16-character password
4. Set environment variables (Windows PowerShell):
   $env:GMAIL_ADDRESS="your-email@gmail.com"
   $env:GMAIL_PASSWORD="your-16-char-app-password"

5. Or add to .env file (create in project root):
   GMAIL_ADDRESS=your-email@gmail.com
   GMAIL_PASSWORD=your-16-char-app-password
"""

import os
from dotenv import load_dotenv

load_dotenv()


def _clean_env(value):
    """Normalize environment values for Gmail configuration."""
    if value is None:
        return ""
    return str(value).strip()


GMAIL_CONFIG = {
    'sender_email': _clean_env(os.getenv('GMAIL_ADDRESS') or os.getenv('EMAIL_ADDRESS')),
    'sender_password': (_clean_env(os.getenv('GMAIL_PASSWORD') or os.getenv('EMAIL_PASSWORD'))).replace(' ', ''),
    'smtp_server': _clean_env(os.getenv('GMAIL_SMTP_SERVER', 'smtp.gmail.com')),
    'smtp_port': int(_clean_env(os.getenv('GMAIL_SMTP_PORT', '465')) or 465),
}


def validate_gmail_config():
    """Check if Gmail is properly configured."""
    if not GMAIL_CONFIG['sender_email'] or not GMAIL_CONFIG['sender_password']:
        return False, "Gmail credentials not configured."
    if GMAIL_CONFIG['smtp_port'] not in (465, 587):
        return False, "Gmail SMTP port must be 465 or 587."
    return True, "Gmail configured successfully"
