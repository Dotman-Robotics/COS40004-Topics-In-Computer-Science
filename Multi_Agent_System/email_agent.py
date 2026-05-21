import json
import win32com.client as win32
from ollama_client import ask_llm
from vector import find_email
from utils import extract_first_json, get_outlook_inbox
from email_summarizer import summarize_email, summarize_inbox_batch, print_email_summary, print_batch_summary


def email_agent(task_details: str | dict, context: str = "") -> dict:
    """Return {to, subject, body} or {error}."""

    if isinstance(task_details, dict) and all(
        k in task_details for k in ("to", "subject", "body")
    ):
        return task_details

    task_str = str(task_details).strip()
    person   = find_email(task_str, interactive=True)
    if not person:
        return {"error": f"No contact confirmed for: {task_str}"}

    email_brief = context.strip() if context.strip() else task_str
    prompt = f"""Write a professional email. Output JSON only. No explanation.

Recipient: {person['name']}
Request: {email_brief}

{{"subject": "...", "body": "..."}}

JSON:"""

    response = ask_llm(prompt)
    cleaned  = extract_first_json(response)

    if not cleaned:
        return {"error": "No JSON in LLM response", "raw": response}

    try:
        email_json = json.loads(cleaned)
    except json.JSONDecodeError as e:
        print(f"[email_agent] JSON parse error: {e}")
        return {"error": "Invalid JSON from LLM", "raw": response}

    email_json["to"] = person["email"]

    for field in ("subject", "body"):
        if not email_json.get(field, "").strip():
            return {"error": f"LLM response missing field: '{field}'", "raw": response}

    return email_json


def send_outlook_email(email_data: dict, sender_account: str | None = None) -> dict | None:
    if "error" in email_data:
        print(f"[email] Skipping send — error: {email_data['error']}")
        return None

    missing = [k for k in ("to", "subject", "body") if not email_data.get(k)]
    if missing:
        print(f"[email] Missing fields: {missing}")
        return None

    try:
        ol_app = win32.Dispatch("Outlook.Application")
        ol_ns  = ol_app.GetNameSpace("MAPI")
        mail   = ol_app.CreateItem(0)

        mail.Subject = email_data["subject"]
        mail.Body    = email_data["body"]
        mail.To      = email_data["to"]

        if sender_account:
            try:
                account = ol_ns.Accounts.Item(sender_account)
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
            except Exception as e:
                print(f"[email] Could not set sender account: {e}")

        mail.Display()
        return email_data

    except Exception as e:
        print(f"[email] Outlook error: {e}")
        return None


def read_outlook_emails(
    account: str = "zoomertron@outlook.com",
    max_emails: int = 5,
    unread_only: bool = True,
    summarize: bool = True,
) -> list[dict]:

    try:
        inbox = get_outlook_inbox(account)
    except ValueError as e:
        print(f"[email] {e}")
        return []

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    emails = []
    for msg in messages:
        try:
            if unread_only and not msg.UnRead:
                continue
            emails.append({
                "from":          msg.SenderName,
                "email":         msg.SenderEmailAddress,
                "subject":       msg.Subject,
                "body":          msg.Body[:500],
                "received_time": str(msg.ReceivedTime),
            })
            if len(emails) >= max_emails:
                break
        except Exception as e:
            print(f"[email] Error reading message: {e}")
            continue

    if summarize and emails:
        print(f"[email] Summarizing {len(emails)} email(s)...")
        for email in emails:
            email["summary"] = summarize_email(email)

    return emails


def display_inbox(emails: list[dict], batch_summary: bool = True) -> None:
    if not emails:
        print("\n  No emails to display.")
        return

    if batch_summary and len(emails) > 1:
        digest = summarize_inbox_batch(emails)
        print_batch_summary(digest)
    else:
        for i, email in enumerate(emails, 1):
            summary = email.get("summary") or summarize_email(email)
            print(f"\n{'─'*50}")
            print(f"  From:     {email['from']} ({email['email']})")
            print(f"  Subject:  {email['subject']}")
            print(f"  Received: {email['received_time']}")
            print_email_summary(summary, index=i)
            print(f"  Body preview: {email['body'][:200]}")


def scan_unread_emails() -> list[dict]:
    try:
        outlook  = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox    = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        emails = []
        for msg in messages:
            try:
                if msg.UnRead:
                    emails.append({
                        "from":    msg.SenderName,
                        "email":   msg.SenderEmailAddress,
                        "subject": msg.Subject,
                        "body":    msg.Body,
                        "time":    str(msg.ReceivedTime),
                    })
            except Exception as e:
                print(f"[email] Skipping message: {e}")
                continue
        return emails

    except Exception as e:
        print(f"[email] Failed to open inbox: {e}")
        return []
    