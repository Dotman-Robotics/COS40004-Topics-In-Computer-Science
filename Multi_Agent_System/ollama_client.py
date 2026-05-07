import requests
 
OLLAMA_URL_GENERATE = "http://localhost:11434/api/generate"
OLLAMA_URL_CHAT     = "http://localhost:11434/api/chat"
MODEL = "llama3.2"
 
 
def ask_llm(prompt: str, model: str = MODEL) -> str: #General Purporse Generation End Point - Email/Calendar 

    try:
        response = requests.post(
            OLLAMA_URL_GENERATE,
            json={
                "model": model,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.1,
                    "num_predict": 512,
                },
            },
            timeout=120,
        )
        response.raise_for_status()
        return response.json()["response"].strip()
    except requests.exceptions.ConnectionError:
        raise RuntimeError("Cannot connect to Ollama. Run 'ollama serve' first.")
    except requests.exceptions.Timeout:
        raise RuntimeError("Ollama timed out.")
    except (KeyError, requests.exceptions.HTTPError) as e:
        raise RuntimeError(f"Ollama error: {e}")
 
 
def ask_llm_chat(system: str, user: str, model: str = MODEL) -> str: #Specialized Generation End Point - Planner/Email Monitor
    """ Note Left for Later Use
    Chat endpoint with a dedicated system prompt.
    Llama 3.2 follows system prompts much more strictly than inline instructions,
    making this significantly more reliable for structured JSON outputs.
 
    Used by the planner agent where consistent JSON formatting is critical.
    """
    try:
        response = requests.post(
            OLLAMA_URL_CHAT,
            json={
                "model": model,
                "stream": False,
                "options": {
                    "temperature": 0.1,
                    "num_predict": 512,
                },
                "messages": [
                    {"role": "system",    "content": system},
                    {"role": "user",      "content": user},
                ],
            },
            timeout=120,
        )
        response.raise_for_status()
        return response.json()["message"]["content"].strip()
    except requests.exceptions.ConnectionError:
        raise RuntimeError("Cannot connect to Ollama. Run 'ollama serve' first.")
    except requests.exceptions.Timeout:
        raise RuntimeError("Ollama timed out.")
    except (KeyError, requests.exceptions.HTTPError) as e:
        raise RuntimeError(f"Ollama error: {e}")
