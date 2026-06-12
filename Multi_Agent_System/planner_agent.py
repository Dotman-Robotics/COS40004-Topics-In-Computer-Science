import json
from ollama_client import ask_llm_chat
from utils import extract_first_json

# RAG WEEK 12: import context query
try:
    from rag_store import query_context
    _RAG_AVAILABLE = True
except ImportError:
    _RAG_AVAILABLE = False

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


def planner_agent(user_input: str, sender_email: str | None = None) -> dict:
    """
    Route a user request to the appropriate agents.

    RAG WEEK 12: if prior context is available for this sender or topic,
    it is prepended to the user message so the LLM can make a more
    informed routing decision. For example, if a previous negotiation
    with a contact failed, the planner can note that context.
    """
    # RAG WEEK 12: retrieve relevant context before routing
    context_prefix = ""
    if _RAG_AVAILABLE:
        try:
            context_prefix = query_context(
                query        = user_input,
                sender_email = sender_email,
                top_k        = 2,
            )
        except Exception:
            pass

    if context_prefix:
        user_message = (
            f"{context_prefix}\n\n"
            f"Request: \"{user_input}\"\nJSON:"
        )
        print(f"[planner] RAG context injected ({len(context_prefix)} chars).")
    else:
        user_message = f'Request: "{user_input}"\nJSON:'

    response = ask_llm_chat(PLANNER_SYSTEM_PROMPT, user_message)
    cleaned  = extract_first_json(response)

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