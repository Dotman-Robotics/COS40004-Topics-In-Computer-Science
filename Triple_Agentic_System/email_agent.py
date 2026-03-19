from ollama_client import ask_llm

def email_agent(task_details):
    prompt = f"""
            Write a professional email based on this request:

            {task_details}

            Return:
            TO:
            SUBJECT:
            BODY:
            """

    return ask_llm(prompt)