"""
Authentication and Authorization Utilities
"""

import secrets
import smtplib
from functools import wraps
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import session, redirect, url_for, flash
from email_config import GMAIL_CONFIG

def require_role(*allowed_roles):
    """Decorator to check user role for route access"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'username' not in session:
                flash('Please log in first', 'danger')
                return redirect(url_for('login'))
            
            user_role = session.get('role', 'user')
            if user_role not in allowed_roles:
                flash('You do not have permission to access this page', 'danger')
                return redirect(url_for('home' if user_role == 'admin' else 'user_home'))
            
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def is_admin():
    """Check if current user is admin"""
    return session.get('role') == 'admin'

def is_user():
    """Check if current user is regular user"""
    return session.get('role') == 'user'

def generate_reset_token():
    """Generate a secure reset token"""
    return secrets.token_urlsafe(32)

def send_password_reset_email(to_email, username, reset_link):
    """
    Send password reset email via Gmail SMTP
    Args:
        to_email: Recipient email
        username: Username for personalization
        reset_link: Full URL for reset link
    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        if not GMAIL_CONFIG['sender_email'] or not GMAIL_CONFIG['sender_password']:
            return False, "Email service not configured. Contact administrator."
        
        # Create email
        msg = MIMEMultipart('alternative')
        msg['Subject'] = 'Barangay Information System - Password Reset Request'
        msg['From'] = GMAIL_CONFIG['sender_email']
        msg['To'] = to_email
        
        # Plain text version
        text = f"""
Hello {username},

You requested to reset your password for the Barangay Information System.
Click the link below to reset your password:

{reset_link}

This link will expire in 1 hour.

If you did not request this, please ignore this email.

Best regards,
Barangay Information System
        """
        
        # HTML version
        html = f"""
        <html>
            <body style="font-family: Arial, sans-serif;">
                <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
                    <h2 style="color: #1B3A57;">Password Reset Request</h2>
                    <p>Hello <strong>{username}</strong>,</p>
                    <p>You requested to reset your password for the Barangay Information System.</p>
                    <p>
                        <a href="{reset_link}" 
                           style="display: inline-block; padding: 12px 24px; background-color: #2e86c1; 
                                  color: white; text-decoration: none; border-radius: 4px; margin: 20px 0;">
                            Reset Your Password
                        </a>
                    </p>
                    <p style="color: #7a93aa; font-size: 12px;">
                        <strong>Link expires in:</strong> 1 hour
                    </p>
                    <p style="color: #7a93aa; font-size: 12px;">
                        If you did not request this, please ignore this email and your password will remain unchanged.
                    </p>
                    <hr style="border: none; border-top: 1px solid #dde3ed; margin: 30px 0;">
                    <p style="color: #7a93aa; font-size: 11px;">
                        Barangay Information System<br>
                        Municipality of Bongabong, Oriental Mindoro
                    </p>
                </div>
            </body>
        </html>
        """
        
        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')
        msg.attach(part1)
        msg.attach(part2)
        
        # Send email
        server = smtplib.SMTP(GMAIL_CONFIG['smtp_server'], GMAIL_CONFIG['smtp_port'])
        server.starttls()
        server.login(GMAIL_CONFIG['sender_email'], GMAIL_CONFIG['sender_password'])
        server.send_message(msg)
        server.quit()
        
        return True, "Password reset email sent successfully"
    
    except smtplib.SMTPAuthenticationError:
        return False, "Email authentication failed. Check Gmail credentials."
    except smtplib.SMTPException as e:
        return False, f"Failed to send email: {str(e)}"
    except Exception as e:
        return False, f"Unexpected error sending email: {str(e)}"

def send_admin_notification(admin_email, username, action):
    """
    Send notification to admin about password reset
    Args:
        admin_email: Admin email address
        username: User who reset password
        action: Action description
    """
    try:
        if not GMAIL_CONFIG['sender_email'] or not GMAIL_CONFIG['sender_password']:
            return False, "Email service not configured"
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = 'Password Reset Activity - Barangay Information System'
        msg['From'] = GMAIL_CONFIG['sender_email']
        msg['To'] = admin_email
        
        text = f"""
Admin Alert,

User '{username}' reset their password.
Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Please review in the audit log if needed.

Barangay Information System
        """
        
        html = f"""
        <html>
            <body style="font-family: Arial, sans-serif;">
                <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
                    <h2 style="color: #1B3A57;">Admin Alert</h2>
                    <p>User <strong>'{username}'</strong> has reset their password.</p>
                    <p>
                        <strong>Timestamp:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br>
                        <strong>Action:</strong> {action}
                    </p>
                    <p>Review the password reset activity in the audit log.</p>
                </div>
            </body>
        </html>
        """
        
        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')
        msg.attach(part1)
        msg.attach(part2)
        
        server = smtplib.SMTP(GMAIL_CONFIG['smtp_server'], GMAIL_CONFIG['smtp_port'])
        server.starttls()
        server.login(GMAIL_CONFIG['sender_email'], GMAIL_CONFIG['sender_password'])
        server.send_message(msg)
        server.quit()
        
        return True, "Admin notification sent"
    except Exception as e:
        print(f"Error sending admin notification: {e}")
        return False, str(e)
