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

GMAIL_CONFIG = {
    'sender_email': os.getenv('GMAIL_ADDRESS', ''),
    'sender_password': os.getenv('GMAIL_PASSWORD', ''),
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587
}

def validate_gmail_config():
    """Check if Gmail is properly configured"""
    if not GMAIL_CONFIG['sender_email'] or not GMAIL_CONFIG['sender_password']:
        return False, "Gmail credentials not configured. See email_config.py for setup instructions."
    return True, "Gmail configured successfully"
