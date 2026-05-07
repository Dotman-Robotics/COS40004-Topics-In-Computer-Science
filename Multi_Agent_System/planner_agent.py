import json
from ollama_client import ask_llm_chat
from vector import find_email


def _extract_first_json(text: str) -> str: #Need to make this its own function
    """Extract the first complete JSON object from text using brace-depth counting."""
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


# System prompt is set once and followed strictly by Llama 3.2.
PLANNER_SYSTEM_PROMPT = """You are a task planner. You output only JSON. Never explain. Never add text outside the JSON.

You decide which actions to take based on the user's request.

Available actions:
- "email"       : send an email. input = recipient name only (e.g. "Johnson")
- "calendar"    : create a calendar event. input = full event description
- "check_inbox" : read emails. input = "". Only use if user asks to read/check/view emails.

A request can require multiple actions. Always include all that apply.

Output format — nothing else:
{"actions": [{"type": "ACTION", "input": "VALUE"}]}"""


def planner_agent(user_input: str) -> dict: #Note left for future reference
    """
    Uses the chat endpoint (ask_llm_chat) instead of generate (ask_llm) because
    Llama 3.2 follows system-role instructions far more reliably than inline
    prompt instructions, especially for compound multi-action requests.
    """
    contact = find_email(user_input)
    contact_hint = (
        f"Contact found: {contact['name']} <{contact['email']}>"
        if contact else
        "No contact found."
    )

    user_message = f"""{contact_hint}
Request: "{user_input}"

JSON:"""

    response = ask_llm_chat(PLANNER_SYSTEM_PROMPT, user_message)
    cleaned  = _extract_first_json(response)

    if not cleaned:
        print(f"[planner] No JSON found in response:\n{response}")
        return {"actions": []}

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        print(f"[planner] JSON parse error: {e}")
        print(f"[planner] Extracted: {cleaned}")
        print(f"[planner] Raw: {response}")
        return {"actions": []}