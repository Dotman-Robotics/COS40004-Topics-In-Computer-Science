from ollama_client import ask_llm


def is_meeting_request(text):
    prompt = f"""
    Does this message contain a meeting request?

    "{text}"

    Answer ONLY: YES or NO
    """

    response = ask_llm(prompt)
    return "YES" in response.upper()