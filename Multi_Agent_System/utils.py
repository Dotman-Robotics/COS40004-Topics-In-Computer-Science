import re
import json


def extract_first_json(text: str) -> str:

    # Strip markdown fences
    text = re.sub(r"```(?:json)?", "", text).strip()

    start = text.find("{")
    if start == -1:
        return ""
    depth = 0
    for i, char in enumerate(text[start:], start):
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start:i + 1]
    return ""


def safe_parse_json(text: str) -> dict | None:

    cleaned = extract_first_json(text)
    if not cleaned:
        return None
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        # Retry with boolean normalisation for LLM quirks
        fixed = re.sub(r'\bTrue\b', 'true', cleaned)
        fixed = re.sub(r'\bFalse\b', 'false', fixed)
        try:
            return json.loads(fixed)
        except json.JSONDecodeError:
            return None


def get_outlook_inbox(email_address: str):

    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for account in outlook.Folders:
        try:
            if account.Name.lower() == email_address.lower():
                return account.Folders["Inbox"]
        except Exception:
            continue
    raise ValueError(f"Outlook account '{email_address}' not found.")