import os
import smtplib
from datetime import datetime, timezone
from email.mime.text import MIMEText


def send_error_report(message: str, context: str = "") -> None:
    """Email an error report via Gmail SMTP using an app password."""
    gmail_address = os.environ.get("GMAIL_ADDRESS", "")
    app_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    if not gmail_address or not app_password:
        raise RuntimeError("GMAIL_ADDRESS / GMAIL_APP_PASSWORD 환경변수가 설정되지 않았습니다.")
    recipient = os.environ.get("NOTIFY_EMAIL", gmail_address)

    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    body = f"Agency Deck Translator에서 오류가 신고되었습니다.\n\n시각: {timestamp}\n\n오류 메시지:\n{message}\n"
    if context:
        body += f"\n상황:\n{context}\n"

    msg = MIMEText(body)
    msg["Subject"] = "[Agency Deck Translator] 오류 신고"
    msg["From"] = gmail_address
    msg["To"] = recipient

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as server:
        server.login(gmail_address, app_password)
        server.sendmail(gmail_address, [recipient], msg.as_string())
