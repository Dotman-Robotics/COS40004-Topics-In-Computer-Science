from ollama_client import ask_llm
from vector import find_email


def email_agent(task_details):
    person = find_email(task_details)

    if not person:
        return f" No matching contact found for: {task_details}"

    prompt = f"""
            Write a professional email based on this request:

            Recipient Name: {person['name']}
            Recipient Email: {person['email']}

            Request:
            {task_details}

            Recipient Email Address:
            Return:
            TO:
            SUBJECT:
            BODY:
            """

    return ask_llm(prompt)