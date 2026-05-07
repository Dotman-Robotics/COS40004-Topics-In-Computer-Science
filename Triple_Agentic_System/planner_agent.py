from ollama_client import ask_llm
from vector import retriver
import json

def planner_agent(user_input):

    email = retriver.invoke(user_input)

    prompt = f"""
You are an AI planner.

Decide which actions to take.

Available actions:
- email
- calendar
- check_inbox

Return ONLY valid JSON like this:

{{
  "actions": [
    {{"type": "email", "input": "..."}},
    {{"type": "calendar", "input": "..."}},
    {{"type": "check_inbox", "input": "..."}}
  ]
}}

Rules:
- Use "check_inbox" if the user asks to read, check, or view emails.
- For email actions, include recipient email from: {email} if available.

if applicable, ensure the input for emails includes recipient email addresses from: {email}, else the recipient email address should just be their name.

User request: {user_input}
"""

    
    response = ask_llm(prompt)

    try:
        return json.loads(response)
    except:
        print("Failed to parse JSON. Raw output:")
        print(response)
        return {"actions": []}