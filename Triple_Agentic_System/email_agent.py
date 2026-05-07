from ollama_client import ask_llm
from vector import find_email
import win32com.client as win32
import win32com.client
import json


def email_agent(task_details):
    if isinstance(task_details, dict):
        return task_details

    task_details_str = str(task_details)

    person = find_email(task_details_str)
    if not person:
        return {"error": f"No matching contact found for: {task_details_str}"}

    prompt = f"""
    Write a professional email.

    Recipient: {person['name']} ({person['email']})

    Request:
    {task_details_str}

    Return ONLY JSON:
    {{
        "to": "...",
        "subject": "...",
        "body": "..."
    }}
    """

    response_text = ask_llm(prompt)

    try:
        email_json = json.loads(response_text)
    except:
        return {"error": "Invalid JSON", "raw": response_text}

    return email_json


def send_outlook_email(email_data, sender_account=None):
    if "error" in email_data:
        print("Error:", email_data)
        return

    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    mailItem = olApp.CreateItem(0)
    mailItem.Subject = email_data["subject"]
    mailItem.Body = email_data["body"]
    mailItem.To = email_data["to"]

    if sender_account:
        mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item(sender_account)))

    mailItem.Display()
    return email_data


def get_inbox_by_account(email_address):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    for account in outlook.Folders:
        try:
            if account.Name.lower() == email_address.lower():
                return account.Folders["Inbox"]
        except:
            continue


def read_outlook_emails(max_emails=1, unread_only=True):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = get_inbox_by_account("lateralus.lateralus.40004@outlook.com")

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    emails = []
    count = 0

    for msg in messages:
        try:
            if unread_only and not msg.UnRead:
                continue

            emails.append({
                "from": msg.SenderName,
                "email": msg.SenderEmailAddress,
                "subject": msg.Subject,
                "body": msg.Body[:300],
                "received_time": str(msg.ReceivedTime)
            })

            count += 1
            if count >= max_emails:
                return emails

        except Exception as e:
            print("Error reading email:", e)

def scan_unread_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    unread_emails = []

    for msg in messages:
        try:
            if msg.UnRead:
                unread_emails.append({
                    "from": msg.SenderName,
                    "email": msg.SenderEmailAddress,
                    "subject": msg.Subject,
                    "body": msg.Body,
                    "time": str(msg.ReceivedTime)
                })
        except:
            continue

    return unread_emails