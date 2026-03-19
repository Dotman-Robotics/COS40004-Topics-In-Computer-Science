from ollama_client import ask_llm
import json

def planner_agent(user_input):
    prompt = f"""
You are an AI planner.

Decide which actions to take.

Available actions:
- email
- calendar

Return ONLY valid JSON like this:

{{
  "actions": [
    {{"type": "email", "input": "..."}},
    {{"type": "calendar", "input": "..."}}
  ]
}}

User request: {user_input}
"""

    response = ask_llm(prompt)

    try:
        return json.loads(response)
    except:
        print("Failed to parse JSON. Raw output:")
        print(response)
        return {"actions": []}