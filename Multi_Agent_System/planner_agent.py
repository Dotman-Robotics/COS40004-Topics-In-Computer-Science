import json
from ollama_client import ask_llm_chat
from utils import extract_first_json

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


def planner_agent(user_input: str) -> dict:
    user_message = f'Request: "{user_input}"\nJSON:'
    response     = ask_llm_chat(PLANNER_SYSTEM_PROMPT, user_message)
    cleaned      = extract_first_json(response)
 
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
