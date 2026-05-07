from ollama_client import ask_llm
from datetime import datetime
import re
import win32com.client


def parse_event_details(llm_output):
    """
    Extract structured fields from LLM response.
    """
    def extract(field):
        match = re.search(rf"{field}:\s*(.*)", llm_output)
        return match.group(1).strip() if match else None

    return {
        "title": extract("TITLE"),
        "date": extract("DATE"),
        "time": extract("TIME"),
    }


def calendar_agent(task_details):
    prompt = f"""
Extract event details from the text below.

Return in this format ONLY:
TITLE: ...
DATE: YYYY-MM-DD
TIME: HH:MM (24-hour)

Text:
{task_details}
"""

    llm_output = ask_llm(prompt)
    event = parse_event_details(llm_output)

    title = event["title"] or "Untitled Event"
    date = event["date"] or datetime.now().strftime("%Y-%m-%d")
    time = event["time"] or "09:00"

    # Convert to Outlook datetime format
    start_datetime = datetime.strptime(f"{date} {time}", "%Y-%m-%d %H:%M")

    # =========================
    # 1. CREATE .ICS FILE
    # =========================
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    dtstart = start_datetime.strftime("%Y%m%dT%H%M%S")

    ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Agent System//EN
BEGIN:VEVENT
UID:{dtstamp}
DTSTAMP:{dtstamp}
SUMMARY:{title}
DTSTART:{dtstart}
END:VEVENT
END:VCALENDAR
"""

    filename = "event.ics"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(ics_content)

    # =========================
    # 2. ADD TO OUTLOOK (win32com)
    # =========================
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        appointment = outlook.CreateItem(1)  # 1 = AppointmentItem

        appointment.Subject = title
        appointment.Start = start_datetime
        appointment.Duration = 60  # default 1 hour
        appointment.Body = task_details
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15
        appointment.Save()

        outlook_result = "Event added to Outlook Calendar"
    except Exception as e:
        outlook_result = f"Outlook error: {str(e)}"

    return {
        "llm_output": llm_output,
        "parsed_event": event,
        "ics_file": filename,
        "outlook_status": outlook_result
    }