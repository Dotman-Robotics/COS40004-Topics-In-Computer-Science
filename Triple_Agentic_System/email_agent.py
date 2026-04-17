from ollama_client import ask_llm
from vector import find_email
import win32com.client as win32
import json

def email_agent(task_details):
    """
    Generates a professional email using an LLM based on task_details.
    Returns a dictionary: {"to": ..., "subject": ..., "body": ...}
    """
    if isinstance(task_details, dict):
        task_details_str = " ".join(str(v) for v in task_details.values())
    elif isinstance(task_details, list):
        task_details_str = " ".join(str(v) for v in task_details)
    else:
        task_details_str = str(task_details)

    person = find_email(task_details_str)
    if not person:
        return {"error": f"No matching contact found for: {task_details_str}"}

    prompt = f"""
    Write a professional email based on the following request:

    Recipient Name: {person['name']}
    Recipient Email: {person['email']}

    Request Details:
    {task_details_str}

    Return the email as a JSON object with these keys:
    {{
        "to": "<Recipient Email Address>",
        "subject": "<Email Subject>",
        "body": "<Email Body>"
    }}

    Instructions:
    - "to" must be the recipient's email.
    - "subject" must be a concise, professional subject line.
    - "body" must include a proper greeting and a professional message.
    - Return only valid JSON; do not include extra text outside the JSON object.

    Return ONLY a JSON object.
    Do NOT include any text like "Here is the email..." or explanations.
    """

    # Call LLM
    response_text = ask_llm(prompt)

    try:
        email_json = json.loads(response_text)
        if not all(k in email_json for k in ("to", "subject", "body")):
            raise ValueError("Missing required keys in LLM response")
    except Exception as e:
        return {"error": f"Failed to parse LLM response: {e}", "raw_response": response_text}

    return email_json

def send_outlook_email(task_details, sender_account=None):
    email_data = email_agent(task_details)

    if "error" in email_data:
        print("Error generating email:", email_data)
        return

    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    mailItem = olApp.CreateItem(0)  # 0 = Mail item
    mailItem.Subject = email_data["subject"]
    mailItem.BodyFormat = 1  # 1 = plain text
    mailItem.Body = email_data["body"]
    mailItem.To = email_data["to"]
    mailItem.Sensitivity = 2  # 2 = confidential

    if sender_account:
        mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item(sender_account)))

    mailItem.Display()
    # mailItem.Send() 

    print(f"Email prepared to: {email_data['to']} with subject: {email_data['subject']}")
    return email_data

def get_inbox_by_account(email_address):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    for account in outlook.Folders:
        try:
            if account.Name.lower() == email_address.lower():
                return account.Folders["Inbox"]
        except:
            continue

def read_outlook_emails(max_emails=5, unread_only=True):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = get_inbox_by_account("lateralus.lateralus.40004@outlook.com")  # Inbox

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
                "body": msg.Body[:300],  # truncate
                "received_time": str(msg.ReceivedTime)
            })
            print(msg.SenderEmailAddress)
            count += 1
            print(count)
            if count >= max_emails:
                return emails
        
        except Exception as e:
            print("Error reading email:", e)
