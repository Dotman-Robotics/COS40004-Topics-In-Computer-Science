import imaplib
import email


IMAP_SERVER = "outlook.office365.com"
EMAIL_ACCOUNT = "lateralus.lateralus.40004@outlook.com"
PASSWORD = "Lateralus@40004"


def get_unread_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    mail.select("inbox")

    status, messages = mail.search(None, '(UNSEEN)')
    email_ids = messages[0].split()

    emails = []

    for e_id in email_ids[-5:]:
        status, msg_data = mail.fetch(e_id, "(RFC822)")
        raw_email = msg_data[0][1]

        msg = email.message_from_bytes(raw_email)

        subject = msg["subject"]
        from_ = msg["from"]

        # Extract body
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode()
                    break
        else:
            body = msg.get_payload(decode=True).decode()

        emails.append({
            "id": e_id,
            "subject": subject,
            "from": from_,
            "body": body
        })

    return emails, mail