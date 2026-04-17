import smtplib
from email.mime.text import MIMEText

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

EMAIL_ACCOUNT = "lateralus.lateralus.40004@outlook.com"
PASSWORD = "Lateralus@40004"


def send_email(to_email, subject, body):
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = EMAIL_ACCOUNT
    msg["To"] = to_email

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ACCOUNT, PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print("SMTP Error:", e)
        return False