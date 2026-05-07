import time
import threading
import json
from datetime import datetime
from ollama_client import ask_llm_chat
from calendar_agent import check_availability, calendar_agent
from email_summarizer import (
    summarize_email,
    summarize_monitor_session,
    print_email_summary,
    print_session_summary,
)

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("[monitor] win32com not available.")

POLL_INTERVAL_SECONDS = 30
MONITOR_ACCOUNT       = "zoomertron@outlook.com"
SENDER_ACCOUNT        = "zoomertron@outlook.com"
_monitor_thread: threading.Thread | None = None
_stop_event     = threading.Event()
_session_log: list[dict] = []
_session_log_lock        = threading.Lock() #post session email processing

def _extract_first_json(text: str) -> str: #seriously i need a function here
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

# May require slight tweaking
DETECTOR_SYSTEM = """You are an email classifier. Output JSON only. Never explain.

Determine if the email is a meeting request. Extract details if it is.

IMPORTANT: Only populate "date" and "time" if they are EXPLICITLY stated in the email.
If the date or time must be inferred or guessed, leave those fields as "".

Output:
{"is_meeting_request": true/false, "title": "...", "date": "YYYY-MM-DD or empty", "time": "HH:MM or empty", "duration_minutes": 60, "requester_name": "...", "requester_email": "..."}

If not a meeting request: {"is_meeting_request": false}"""


def detect_meeting_request(email: dict) -> dict:
    """Use the LLM to determine if an email is a meeting request."""
    today    = datetime.now().strftime("%A %d %B %Y")
    user_msg = f"""Today is {today}.

From: {email['from']} <{email['email']}>
Subject: {email['subject']}
Body: {email['body']}

JSON:"""

    response = ask_llm_chat(DETECTOR_SYSTEM, user_msg)
    cleaned  = _extract_first_json(response)

    if not cleaned:
        return {"is_meeting_request": False}

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        return {"is_meeting_request": False}

# Reply Composition may need to be more specific
REPLY_SYSTEM = """You are an email assistant. Output JSON only. Never explain.
Write a professional reply email.
Output: {"subject": "...", "body": "..."}"""

def _compose_reply(
    original_email: dict,
    meeting: dict,
    confirmed: bool,
    alternatives: list[str] | None = None,
) -> dict:
    if confirmed:
        instruction = (
            f"Write a confirmation that the meeting '{meeting.get('title', 'Meeting')}' "
            f"on {meeting.get('date')} at {meeting.get('time')} has been booked."
        )
    else:
        alt_text = ""
        if alternatives:
            alt_text = f"Suggest these alternative times on {meeting.get('date')}: {', '.join(alternatives)}."
        instruction = (
            f"Write a polite decline for the meeting '{meeting.get('title', 'Meeting')}' "
            f"on {meeting.get('date')} at {meeting.get('time')} due to a scheduling conflict. "
            f"{alt_text}"
        )

    user_msg = f"""Reply to: {original_email['from']} <{original_email['email']}>
Original subject: {original_email['subject']}
Instruction: {instruction}

JSON:"""

    response = ask_llm_chat(REPLY_SYSTEM, user_msg)
    cleaned  = _extract_first_json(response)

    try:
        reply = json.loads(cleaned) if cleaned else {}
    except json.JSONDecodeError:
        reply = {}

    status_word = "Confirmed" if confirmed else "Declined"
    reply.setdefault("subject", f"Re: {original_email['subject']} — {status_word}")
    reply.setdefault("body",    "Meeting confirmed." if confirmed else "Unable to accommodate — please suggest an alternative.")
    reply["to"] = original_email["email"]

    return reply

def _get_inbox(): # To be merged with another function from email agent
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for account in outlook.Folders:
        try:
            if account.Name.lower() == MONITOR_ACCOUNT.lower():
                return account.Folders["Inbox"]
        except Exception:
            continue
    raise ValueError(f"Account '{MONITOR_ACCOUNT}' not found in Outlook.")


def _send_reply(reply: dict) -> bool:
    try:
        ol_app = win32com.client.Dispatch("Outlook.Application")
        ol_ns  = ol_app.GetNameSpace("MAPI")
        mail   = ol_app.CreateItem(0)

        mail.Subject = reply["subject"]
        mail.Body    = reply["body"]
        mail.To      = reply["to"]

        try:
            account = ol_ns.Accounts.Item(SENDER_ACCOUNT)
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
        except Exception as e:
            print(f"[monitor] Could not set sender account: {e}")

        mail.Display()  # No point in this I already triggered spam limits lmao
        return True
    except Exception as e:
        print(f"[monitor] Failed to send reply: {e}")
        return False


def _mark_as_read(msg) -> None:
    try:
        msg.UnRead = False
        msg.Save()
    except Exception as e:
        print(f"[monitor] Could not mark as read: {e}")

def process_email(msg) -> dict | None: #logs all meetings, SHOULD summarize them at the end
    try:
        email = {
            "from":          msg.SenderName,
            "email":         msg.SenderEmailAddress,
            "subject":       msg.Subject,
            "body":          msg.Body[:500],
            "received_time": str(msg.ReceivedTime),
        }
    except Exception as e:
        print(f"[monitor] Could not read message: {e}")
        return None

    print(f"\n[monitor] ── New email from {email['from']}: '{email['subject']}'")
    summary = summarize_email(email)
    print(f"[monitor] Summary: {summary.get('summary', '')}")
    print(f"[monitor] Category: {summary.get('category', 'other')} | Priority: {summary.get('priority', 'medium')}")
    if summary.get("action_items"):
        print(f"[monitor] Action items: {summary['action_items']}")

    email["summary"] = summary

    meeting = detect_meeting_request(email)

    if not meeting.get("is_meeting_request"):
        print(f"[monitor] Not a meeting request — logged and skipped.") #SKIP
        _mark_as_read(msg)
        result = {
            "status":  "non_meeting",
            "email":   email,
            "meeting": None,
            "summary": summary,
        }
        with _session_log_lock:
            _session_log.append(result)
        return result

    print(f"[monitor] Meeting request detected: {meeting.get('title')} on {meeting.get('date')} at {meeting.get('time')}")

    date = meeting.get("date", "").strip()
    time = meeting.get("time", "").strip()

    if not date or not time: #Request Date and Time
        print("[monitor] Missing explicit date or time — requesting clarification.")
        clarification = {
            "to":      email["email"],
            "subject": f"Re: {email['subject']} — Time Needed",
            "body": (
                f"Hi {email['from']},\n\n"
                f"Thank you for your message about '{meeting.get('title', 'the meeting')}'. "
                f"Could you please specify a preferred date and time? "
                f"I'll check availability and confirm shortly.\n\nBest regards"
            ),
        }
        sent = _send_reply(clarification)
        print(f"[monitor] Clarification request sent: {sent}")
        _mark_as_read(msg)
        result = {
            "status":  "clarification_needed",
            "email":   email,
            "meeting": meeting,
            "summary": summary,
        }
        with _session_log_lock:
            _session_log.append(result)
        return result

    body_lower = email["body"].lower()
    time_indicators = ["am", "pm", "o'clock", "noon", "midnight", "morning",
                       "afternoon", "evening", "night", "at ", ":", "hour"]
    time_was_explicit = any(ind in body_lower for ind in time_indicators)

    if not time_was_explicit: # GET TIME
        print(f"[monitor] Time not explicit — requesting clarification.")
        clarification = {
            "to":      email["email"],
            "subject": f"Re: {email['subject']} — Time Needed",
            "body": (
                f"Hi {email['from']},\n\n"
                f"Thank you for reaching out about '{meeting.get('title', 'the meeting')}'. "
                f"Could you please confirm the preferred time? "
                f"I'll check availability and get back to you.\n\nBest regards"
            ),
        }
        sent = _send_reply(clarification)
        _mark_as_read(msg)
        result = {
            "status":  "clarification_needed",
            "email":   email,
            "meeting": meeting,
            "summary": summary,
        }
        with _session_log_lock:
            _session_log.append(result)
        return result

    duration = int(meeting.get("duration_minutes") or 60)
    availability = check_availability(date, time, duration) #Check Avails
    print(f"[monitor] Availability: {'Free' if availability['available'] else 'Conflict'}")

    if availability["available"]:
        event_desc = (
            f"{meeting.get('title', 'Meeting')} with {email['from']} "
            f"on {date} at {time} for {duration} minutes"
        )
        cal_result = calendar_agent(event_desc)
        print(f"[monitor] Booked: {cal_result.get('outlook_status')}")

        reply = _compose_reply(email, meeting, confirmed=True)
        sent  = _send_reply(reply)
        print(f"[monitor] Confirmation reply sent: {sent}")

        _mark_as_read(msg)
        result = {
            "status":    "confirmed",
            "email":     email,
            "meeting":   meeting,
            "calendar":  cal_result,
            "reply_sent": sent,
            "summary":   summary,
        }

    else: #Conflicts
        alternatives = availability.get("suggested_alternatives", [])
        print(f"[monitor] Conflict — alternatives: {alternatives}")

        reply = _compose_reply(email, meeting, confirmed=False, alternatives=alternatives)
        sent  = _send_reply(reply)
        print(f"[monitor] Cancellation reply sent: {sent}")

        _mark_as_read(msg)
        result = {
            "status":       "declined",
            "email":        email,
            "meeting":      meeting,
            "conflicts":    availability["conflicts"],
            "alternatives": alternatives,
            "reply_sent":   sent,
            "summary":      summary,
        }

    with _session_log_lock:
        _session_log.append(result)

    return result


def poll_inbox() -> list[dict]:
    """Check the inbox once and process all unread emails."""
    if not WIN32_AVAILABLE:
        print("[monitor] win32com not available.")
        return []

    results = []
    try:
        inbox    = _get_inbox()
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        unread   = [msg for msg in messages if msg.UnRead]
        print(f"[monitor] Found {len(unread)} unread email(s).")

        for msg in unread:
            result = process_email(msg)
            if result:
                results.append(result)

    except Exception as e:
        print(f"[monitor] Inbox poll error: {e}")

    return results


# ── Monitor loop ──────────────────────────────────────────────────────────────

def _monitor_loop():
    """Background thread that polls the inbox on a timer."""
    pythoncom.CoInitialize()
    print(f"[monitor] Started. Polling every {POLL_INTERVAL_SECONDS}s.")

    try:
        while not _stop_event.is_set():
            try:
                results = poll_inbox()
                if results:
                    print(f"[monitor] Processed {len(results)} email(s) this poll.")
            except Exception as e:
                print(f"[monitor] Poll error: {e}")
            _stop_event.wait(timeout=POLL_INTERVAL_SECONDS)
    finally:
        pythoncom.CoUninitialize()
        print("[monitor] Thread stopped.")


def start_monitor():
    """Start the inbox monitor in a background thread."""
    global _monitor_thread
    if _monitor_thread and _monitor_thread.is_alive():
        print("[monitor] Already running.")
        return

    # Clear session log for new session
    with _session_log_lock:
        _session_log.clear()

    _stop_event.clear()
    _monitor_thread = threading.Thread(
        target=_monitor_loop, daemon=True, name="EmailMonitor"
    )
    _monitor_thread.start()
    print("[monitor] Monitor started.")


def stop_monitor() -> list[dict]:
    """
    Stop the monitor thread and return the session log.
    The caller (main.py) is responsible for generating and displaying
    the session summary from the returned log.
    """
    global _monitor_thread
    _stop_event.set()
    if _monitor_thread:
        _monitor_thread.join(timeout=5)
    _monitor_thread = None

    with _session_log_lock:
        log_copy = list(_session_log)

    return log_copy


def get_session_log() -> list[dict]:
    """Return a copy of the current session log without stopping the monitor."""
    with _session_log_lock:
        return list(_session_log)


def is_running() -> bool:
    return _monitor_thread is not None and _monitor_thread.is_alive()
