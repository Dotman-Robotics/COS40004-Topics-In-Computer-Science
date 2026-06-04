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

import os

DEFAULT_DURATION = 60
ICS_DIR          = "."

# Outlook BusyStatus constants
OL_TENTATIVE = 1
OL_BUSY      = 2


def _get_account() -> str:
    """
    Read the active account at call time rather than import time.
    This ensures --account CLI arg set by main.py is always respected,
    even if calendar_agent was imported before the env var was written.
    """
    return os.environ.get("AGENT_ACCOUNT", "zoomertron@outlook.com")


def _extract_first_json(text: str) -> str:
    import re as _re
    text  = _re.sub(r"```(?:json)?", "", text).strip()
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
    today  = datetime.now().strftime("%A %d %B %Y")
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
        print(f"[calendar] JSON parse error: {e}")
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

    return {
        "title":            title,
        "date":             date,
        "time":             time,
        "duration_minutes": duration,
        "location":         str(parsed.get("location")    or "").strip(),
        "description":      str(parsed.get("description") or task_details).strip(),
    }


def _build_ics(event: dict, start_dt: datetime, end_dt: datetime) -> str:
    def esc(s):
        return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")

    uid     = str(uuid.uuid4())
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    dtstart = start_dt.strftime("%Y%m%dT%H%M%S")
    dtend   = end_dt.strftime("%Y%m%dT%H%M%S")

    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        "PRODID:-//Agent System//EN", "CALSCALE:GREGORIAN",
        "BEGIN:VEVENT",
        f"UID:{uid}", f"DTSTAMP:{dtstamp}",
        f"DTSTART:{dtstart}", f"DTEND:{dtend}",
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


def _get_account_calendar(account_email: str):
    """Return the Calendar folder for the specified Outlook account."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for folder in outlook.Folders:
        if folder.Name.lower() == account_email.lower():
            try:
                return folder.Folders["Calendar"]
            except Exception:
                for sub in folder.Folders:
                    if sub.DefaultItemType == 1:
                        return sub
    raise ValueError(f"[calendar] Calendar folder not found for '{account_email}'.")


# ── Tentative booking ─────────────────────────────────────────────────────────

def book_tentative(event: dict, start_dt: datetime, end_dt: datetime) -> dict:
    """
    Create a TENTATIVE Outlook appointment (BusyStatus = olTentative).

    Returns:
        {
            "entry_id":      str,   # Outlook EntryID — store this to confirm/cancel later
            "outlook_status": str,
        }

    The slot blocks the calendar so no other event is booked over it, but the
    appointment is visually marked as tentative in Outlook. Promote it to
    confirmed with confirm_tentative(), or remove it with cancel_tentative().
    """
    if not WIN32_AVAILABLE:
        return {"entry_id": None, "outlook_status": "Outlook unavailable."}
    try:
        calendar_folder = _get_account_calendar(_get_account())
        appointment     = calendar_folder.Items.Add(1)

        appointment.Subject    = f"[TENTATIVE] {event['title']}"
        appointment.Start      = start_dt.strftime("%Y-%m-%d %H:%M")
        appointment.End        = end_dt.strftime("%Y-%m-%d %H:%M")
        appointment.Duration   = event["duration_minutes"]
        appointment.Body       = (
            f"{event['description']}\n\n"
            f"[This appointment is TENTATIVE pending confirmation from the other party.]"
        )
        appointment.BusyStatus    = OL_TENTATIVE
        appointment.ReminderSet   = False  # No reminder until confirmed

        if event["location"]:
            appointment.Location = event["location"]

        appointment.Save()
        entry_id = appointment.EntryID
        print(f"[calendar] Tentative booking created: '{event['title']}' on "
              f"{start_dt.strftime('%Y-%m-%d %H:%M')} (EntryID: {entry_id[:16]}…)")

        return {
            "entry_id":       entry_id,
            "outlook_status": f"[TENTATIVE] '{event['title']}' blocked pending confirmation.",
        }

    except ValueError as e:
        return {"entry_id": None, "outlook_status": str(e)}
    except Exception as e:
        return {"entry_id": None, "outlook_status": f"Outlook error: {e}"}


def confirm_tentative(entry_id: str, event_title: str) -> str:
    """
    Promote a tentative appointment to CONFIRMED (BusyStatus = olBusy).
    Removes the [TENTATIVE] prefix from the subject.
    Returns a status string.
    """
    if not WIN32_AVAILABLE:
        return "Outlook unavailable."
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        appointment = outlook.GetItemFromID(entry_id)

        appointment.Subject    = event_title  # strip [TENTATIVE] prefix
        appointment.BusyStatus = OL_BUSY
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15
        appointment.Body = appointment.Body.replace(
            "\n\n[This appointment is TENTATIVE pending confirmation from the other party.]", ""
        ).strip()
        appointment.Save()

        print(f"[calendar] Confirmed: '{event_title}' (EntryID: {entry_id[:16]}…)")
        return f"'{event_title}' confirmed in calendar."

    except Exception as e:
        print(f"[calendar] confirm_tentative error: {e}")
        return f"Could not confirm appointment: {e}"


def cancel_tentative(entry_id: str, event_title: str) -> str:
    """
    Delete a tentative appointment from the calendar.
    Called when negotiation fails, times out, or the other party rejects.
    Returns a status string.
    """
    if not WIN32_AVAILABLE:
        return "Outlook unavailable."
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        appointment = outlook.GetItemFromID(entry_id)
        appointment.Delete()

        print(f"[calendar] Cancelled tentative: '{event_title}' (EntryID: {entry_id[:16]}…)")
        return f"Tentative '{event_title}' removed from calendar."

    except Exception as e:
        print(f"[calendar] cancel_tentative error: {e}")
        return f"Could not cancel tentative appointment: {e}"


# ── Confirmed booking (used by calendar_agent for non-negotiation bookings) ───

def _add_to_outlook(event: dict, start_dt: datetime, end_dt: datetime) -> str:
    """
    Create a fully confirmed Outlook appointment directly.
    Used by calendar_agent() for user-initiated bookings that don't
    go through negotiation (no tentative phase needed).
    """
    if not WIN32_AVAILABLE:
        return "Outlook unavailable."
    try:
        calendar_folder = _get_account_calendar(_get_account())
        appointment     = calendar_folder.Items.Add(1)

        appointment.Subject    = event["title"]
        appointment.Start      = start_dt.strftime("%Y-%m-%d %H:%M")
        appointment.End        = end_dt.strftime("%Y-%m-%d %H:%M")
        appointment.Duration   = event["duration_minutes"]
        appointment.Body       = event["description"]
        appointment.BusyStatus = OL_BUSY
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15

        if event["location"]:
            appointment.Location = event["location"]

        appointment.Save()
        return f"'{event['title']}' added to calendar for {_get_account()}."

    except ValueError as e:
        return str(e)
    except Exception as e:
        return f"Outlook error: {e}"


# ── Busy slot reader ──────────────────────────────────────────────────────────

def get_busy_slots(date_str: str) -> list[dict]:
    """
    Return all appointments (confirmed AND tentative) for date_str.
    Tentative slots are included so the negotiation agent doesn't
    double-book over them.
    """
    if not WIN32_AVAILABLE:
        return []
    try:
        account         = _get_account()
        calendar_folder = _get_account_calendar(account)
        print(f"[calendar] get_busy_slots: checking {account} calendar for {date_str}")

        target    = datetime.strptime(date_str, "%Y-%m-%d")
        day_start = target.replace(hour=0,  minute=0,  second=0)
        day_end   = target.replace(hour=23, minute=59, second=59)

        items = calendar_folder.Items
        items.IncludeRecurrences = True  # must be set BEFORE Sort
        items.Sort("[Start]")

        # Use ISO 8601 format for the restriction — locale-independent
        restriction = (
            f"[Start] >= '{day_start.strftime('%Y-%m-%d %H:%M')}' AND "
            f"[Start] <= '{day_end.strftime('%Y-%m-%d %H:%M')}'"
        )
        restricted = items.Restrict(restriction)

        slots = []
        for item in restricted:
            try:
                start    = item.Start
                end      = item.End
                start_dt = datetime(start.year, start.month, start.day,
                                    start.hour, start.minute)
                end_dt   = datetime(end.year,   end.month,   end.day,
                                    end.hour,   end.minute)
                slots.append({
                    "title":            item.Subject,
                    "start":            start_dt.strftime("%Y-%m-%d %H:%M"),
                    "end":              end_dt.strftime("%Y-%m-%d %H:%M"),
                    "duration_minutes": int((end_dt - start_dt).total_seconds() / 60),
                    "tentative":        item.BusyStatus == OL_TENTATIVE,
                })
            except Exception as e:
                print(f"[calendar] Error reading appointment: {e}")
                continue

        # Fallback: if restriction returned nothing, scan all items manually.
        # Outlook's restriction engine can silently fail on some locale/timezone
        # configurations — iterating manually is slower but always correct.
        if not slots:
            print(f"[calendar] Restriction returned 0 — falling back to full scan.")
            for item in calendar_folder.Items:
                try:
                    start    = item.Start
                    start_dt = datetime(start.year, start.month, start.day,
                                        start.hour, start.minute)
                    if not (day_start <= start_dt <= day_end):
                        continue
                    end      = item.End
                    end_dt   = datetime(end.year, end.month, end.day,
                                        end.hour, end.minute)
                    slots.append({
                        "title":            item.Subject,
                        "start":            start_dt.strftime("%Y-%m-%d %H:%M"),
                        "end":              end_dt.strftime("%Y-%m-%d %H:%M"),
                        "duration_minutes": int((end_dt - start_dt).total_seconds() / 60),
                        "tentative":        item.BusyStatus == OL_TENTATIVE,
                    })
                except Exception:
                    continue

        print(f"[calendar] get_busy_slots: found {len(slots)} appointment(s): "
              f"{[s['title'] + ' ' + s['start'] for s in slots]}")
        return slots

    except Exception as e:
        print(f"[calendar] get_busy_slots error: {e}")
        return []


# Business hours window used for slot search
BUSINESS_HOURS_START = 8   # 08:00 inclusive
BUSINESS_HOURS_END   = 17  # 17:00 exclusive — no slot starting at or after this hour


def check_availability(date_str: str, time_str: str, duration_minutes: int = 60) -> dict:
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

    # Search for free slots within business hours, rolling forward across days
    # if the current day is fully booked.
    alternatives = []
    search_date  = proposed_start.date()
    MAX_DAYS     = 7  # don't search more than a week ahead

    for day_offset in range(MAX_DAYS):
        current_date      = search_date + timedelta(days=day_offset)
        current_date_str  = current_date.strftime("%Y-%m-%d")
        day_busy          = get_busy_slots(current_date_str) if day_offset > 0 else busy_slots

        for hour in range(BUSINESS_HOURS_START, BUSINESS_HOURS_END):
            candidate_start = datetime(
                current_date.year, current_date.month, current_date.day,
                hour, 0
            )
            candidate_end = candidate_start + timedelta(minutes=duration_minutes)

            # Skip the originally proposed slot
            if candidate_start == proposed_start:
                continue

            # Ensure the slot ends within business hours
            if candidate_end.hour > BUSINESS_HOURS_END or (
                candidate_end.hour == BUSINESS_HOURS_END and candidate_end.minute > 0
            ):
                continue

            overlap = any(
                candidate_start < datetime.strptime(s["end"],   "%Y-%m-%d %H:%M") and
                candidate_end   > datetime.strptime(s["start"], "%Y-%m-%d %H:%M")
                for s in day_busy
            )
            if not overlap:
                # Return as "YYYY-MM-DD HH:MM" so the negotiation agent knows the date
                alternatives.append(candidate_start.strftime("%Y-%m-%d %H:%M"))

            if len(alternatives) >= 3:
                break

        if len(alternatives) >= 3:
            break

    return {
        "available":              False,
        "conflicts":              conflicts,
        "suggested_alternatives": alternatives,
    }


def calendar_agent(task_details: str) -> dict:
    """
    For user-initiated bookings (not via negotiation): parse, validate, and
    book directly as confirmed. No tentative phase.
    """
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
        "llm_output":     parsed_raw,
        "parsed_event":   event,
        "availability":   availability,
        "ics_file":       ics_filename,
        "outlook_status": outlook_status,
    }