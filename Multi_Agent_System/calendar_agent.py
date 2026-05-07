import json
import uuid
import re
from datetime import datetime, timedelta
from ollama_client import ask_llm

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("[calendar] win32com not available — Outlook integration disabled.")

DEFAULT_DURATION = 60  # minutes
ICS_DIR = "."

def _extract_first_json(text: str) -> str:
    start = text.find("{")
    if start == -1:
        return ""
    depth = 0
    for i, char in enumerate(text[start:], start):
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start:i + 1]
    return ""

def _parse_event(task_details: str) -> dict:
    today = datetime.now().strftime("%A %d %B %Y")

    prompt = f"""Extract calendar event details. Output JSON only. No explanation.
Today is {today}.

Fields:
- title: event name
- date: YYYY-MM-DD
- time: HH:MM (24hr)
- duration_minutes: integer
- location: string or ""
- description: string or ""

Text: "{task_details}"

JSON:"""

    response = ask_llm(prompt)
    cleaned  = _extract_first_json(response)

    if not cleaned:
        print(f"[calendar] No JSON found:\n{response}") 
        return {}

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        print(f"[calendar] JSON parse error: {e}\nExtracted: {cleaned}")
        return {}


def _apply_defaults(parsed: dict, task_details: str) -> dict:  
    title = str(parsed.get("title") or "Untitled Event").strip()

    try:
        date = datetime.strptime(str(parsed.get("date", "")), "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        date = datetime.now().strftime("%Y-%m-%d")
        print(f"[calendar] Invalid date '{parsed.get('date')}' — using today.")

    try:
        time = datetime.strptime(str(parsed.get("time", "")), "%H:%M").strftime("%H:%M")
    except ValueError:
        time = "09:00"
        print(f"[calendar] Invalid time '{parsed.get('time')}' — defaulting to 09:00.")

    try:
        duration = max(1, int(parsed.get("duration_minutes") or DEFAULT_DURATION))
    except (ValueError, TypeError):
        duration = DEFAULT_DURATION

    location    = str(parsed.get("location")    or "").strip()
    description = str(parsed.get("description") or task_details).strip()

    return {
        "title":            title,
        "date":             date,
        "time":             time,
        "duration_minutes": duration,
        "location":         location,
        "description":      description,
    }

def _build_ics(event: dict, start_dt: datetime, end_dt: datetime) -> str: #RFC 5545 ICS
    def esc(s):
        return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")

    uid     = str(uuid.uuid4())
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    dtstart = start_dt.strftime("%Y%m%dT%H%M%S")
    dtend   = end_dt.strftime("%Y%m%dT%H%M%S")

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Agent System//EN",
        "CALSCALE:GREGORIAN",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{dtstamp}",
        f"DTSTART:{dtstart}",
        f"DTEND:{dtend}",
        f"SUMMARY:{esc(event['title'])}",
    ]
    if event["location"]:
        lines.append(f"LOCATION:{esc(event['location'])}")
    if event["description"]:
        lines.append(f"DESCRIPTION:{esc(event['description'])}")
    lines += ["END:VEVENT", "END:VCALENDAR", ""]

    return "\r\n".join(lines)


def _save_ics(content: str, title: str) -> str:
    safe = re.sub(r"[^\w\s-]", "", title).strip().replace(" ", "_")[:40]
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = f"{ICS_DIR}/event_{safe}_{ts}.ics"
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    return path

def _get_outlook_calendar():
    if not WIN32_AVAILABLE:
        return None
    outlook  = win32com.client.Dispatch("Outlook.Application")
    ns       = outlook.GetNamespace("MAPI")
    return ns.GetDefaultFolder(9)  #9 = Calendar


def _add_to_outlook(event: dict, start_dt: datetime, end_dt: datetime) -> str:

    if not WIN32_AVAILABLE:
        return "Outlook unavailable."
    try:
        outlook     = win32com.client.Dispatch("Outlook.Application")
        appointment = outlook.CreateItem(1)  # 1 =AppointmentItem

        appointment.Subject  = event["title"]
        appointment.Start    = start_dt.strftime("%Y-%m-%d %H:%M")
        appointment.End      = end_dt.strftime("%Y-%m-%d %H:%M")
        appointment.Duration = event["duration_minutes"]
        appointment.Body     = event["description"]
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15

        if event["location"]:
            appointment.Location = event["location"]

        appointment.Save()
        return f"'{event['title']}' added to Outlook Calendar."
    except Exception as e:
        return f"Outlook error: {e}"

def get_busy_slots(date_str: str) -> list[dict]:
    if not WIN32_AVAILABLE:
        return []
    try:
        calendar = _get_outlook_calendar()
        if not calendar:
            return []

        target    = datetime.strptime(date_str, "%Y-%m-%d") #pull specific dates
        day_start = target.replace(hour=0,  minute=0,  second=0)
        day_end   = target.replace(hour=23, minute=59, second=59)

        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        restriction = (
            f"[Start] >= '{day_start.strftime('%m/%d/%Y %H:%M')}' AND "
            f"[Start] <= '{day_end.strftime('%m/%d/%Y %H:%M')}'"
        )
        restricted = items.Restrict(restriction)

        slots = []
        for item in restricted:
            try:
                start = item.Start
                end   = item.End
                #win32com returns pytime objects
                start_dt = datetime(start.year, start.month, start.day, start.hour, start.minute)
                end_dt   = datetime(end.year,   end.month,   end.day,   end.hour,   end.minute)
                slots.append({
                    "title":            item.Subject,
                    "start":            start_dt.strftime("%Y-%m-%d %H:%M"),
                    "end":              end_dt.strftime("%Y-%m-%d %H:%M"),
                    "duration_minutes": int((end_dt - start_dt).total_seconds() / 60),
                })
            except Exception as e:
                print(f"[calendar] Error reading appointment: {e}")
                continue

        return slots

    except Exception as e:
        print(f"[calendar] get_busy_slots error: {e}")
        return []


def check_availability(date_str: str, time_str: str, duration_minutes: int = 60) -> dict: #(date str in YYYY-MM-DD format, time in HH:MM) (Returns available, conflicts and alts)

    try:
        proposed_start = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
        proposed_end   = proposed_start + timedelta(minutes=duration_minutes)
    except ValueError as e:
        return {"available": False, "conflicts": [], "error": str(e)}

    busy_slots = get_busy_slots(date_str)
    conflicts  = []

    for slot in busy_slots:
        slot_start = datetime.strptime(slot["start"], "%Y-%m-%d %H:%M")
        slot_end   = datetime.strptime(slot["end"],   "%Y-%m-%d %H:%M")
        if proposed_start < slot_end and proposed_end > slot_start:
            conflicts.append(slot)

    if not conflicts:
        return {"available": True, "conflicts": [], "suggested_alternatives": []}
    
    alternatives = [] #try every hour from 08:00 to 18:00
    for hour in range(8, 19):
        candidate_start = proposed_start.replace(hour=hour, minute=0)
        candidate_end   = candidate_start + timedelta(minutes=duration_minutes)
        if candidate_start == proposed_start:
            continue

        overlap = False #overlap check
        for slot in busy_slots:
            slot_start = datetime.strptime(slot["start"], "%Y-%m-%d %H:%M")
            slot_end   = datetime.strptime(slot["end"],   "%Y-%m-%d %H:%M")
            if candidate_start < slot_end and candidate_end > slot_start:
                overlap = True
                break

        if not overlap:
            alternatives.append(candidate_start.strftime("%H:%M"))

        if len(alternatives) >= 3:
            break

    return {
        "available":               False,
        "conflicts":               conflicts,
        "suggested_alternatives":  alternatives,
    }

def calendar_agent(task_details: str) -> dict:

    parsed_raw = _parse_event(task_details)
    event      = _apply_defaults(parsed_raw, task_details)
    start_dt   = datetime.strptime(f"{event['date']} {event['time']}", "%Y-%m-%d %H:%M")
    end_dt     = start_dt + timedelta(minutes=event["duration_minutes"])

    availability = check_availability(event["date"], event["time"], event["duration_minutes"])
    if not availability["available"]:
        print(f"[calendar] WARNING: Slot conflicts with: {[c['title'] for c in availability['conflicts']]}")
        if availability.get("suggested_alternatives"):
            print(f"[calendar] Suggested free slots: {availability['suggested_alternatives']}")

    ics_content    = _build_ics(event, start_dt, end_dt)
    ics_filename   = _save_ics(ics_content, event["title"])
    outlook_status = _add_to_outlook(event, start_dt, end_dt)

    return {
        "llm_output":    parsed_raw,
        "parsed_event":  event,
        "availability":  availability,
        "ics_file":      ics_filename,
        "outlook_status": outlook_status,
    }