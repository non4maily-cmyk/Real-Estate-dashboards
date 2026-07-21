"""
وحدة إرسال تنبيه بريدي عبر Gmail SMTP
تُستخدم لإخطار محمد فوراً لو توقف السكريبت لأي سبب
"""
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

def send_alert_email(subject, body_lines, to_addr, from_addr, app_password):
    body = "\n".join(body_lines)
    body += f"\n\n---\nوقت التنفيذ: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    msg = MIMEText(body, 'plain', 'utf-8')
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=30) as server:
            server.starttls()
            server.login(from_addr, app_password)
            server.sendmail(from_addr, [to_addr], msg.as_string())
        return True
    except Exception as e:
        print(f"⚠️ فشل إرسال إيميل التنبيه نفسه: {e}")
        return False
