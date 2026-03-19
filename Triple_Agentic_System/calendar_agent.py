from ollama_client import ask_llm
from datetime import datetime

def calendar_agent(task_details):
    prompt = f"""
            Extract event details:

            {task_details}

            Return:
            TITLE:
            DATE:
            TIME:
            """

    event = ask_llm(prompt)

    # Save as .ics file
    filename = "event.ics"

    with open(filename, "w") as f:
        f.write(f"""BEGIN:VCALENDAR
                    BEGIN:VEVENT
                    SUMMARY:{task_details}
                    DTSTART:20260320T100000
                    END:VEVENT
                    END:VCALENDAR
                    """)

    return f"{event}\n\nSaved to {filename}"