import json
import re
import threading
from datetime import datetime, timedelta

from ollama_client import ask_llm_chat
from calendar_agent import check_availability, calendar_agent, get_busy_slots
from utils import safe_parse_json

NEGOTIATION_STATE_FILE = "negotiations.json"
MAX_ROUNDS             = 5
AGENT_TAG              = "X-AgentSystem: true"
WORK_HOURS_START       = 8
WORK_HOURS_END         = 18

_state_lock = threading.Lock()

def _load_state() -> dict:
    try:
        with open(NEGOTIATION_STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_state(state: dict) -> None:
    with open(NEGOTIATION_STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, default=str)


def _get_negotiation(thread_id: str) -> dict | None:
    with _state_lock:
        return _load_state().get(thread_id)


def _upsert_negotiation(thread_id: str, data: dict) -> None:
    with _state_lock:
        state = _load_state()
        state[thread_id] = data
        _save_state(state)


def _close_negotiation(thread_id: str, final_status: str) -> None:
    with _state_lock:
        state = _load_state()
        if thread_id in state:
            state[thread_id]["status"]    = final_status
            state[thread_id]["closed_at"] = datetime.now().isoformat()
            _save_state(state)


def get_all_negotiations() -> dict:
    with _state_lock:
        return _load_state()


def _extract_thread_id(subject: str) -> str:
    clean = re.sub(r"^(Re|Fwd|FWD|RE|FW):\s*", "", subject, flags=re.IGNORECASE).strip()
    return re.sub(r"\s+", "_", clean.lower())[:80]

def is_agent_email(body: str) -> bool:
    return AGENT_TAG.lower() in body.lower()

def inject_agent_tag(body: str) -> str:
    return body + f"\n\n--\n{AGENT_TAG}"

def get_all_free_slots(date_str: str, duration_minutes: int = 60) -> list[str]:
    busy       = get_busy_slots(date_str)
    free       = []
    slot_start = datetime.strptime(date_str, "%Y-%m-%d").replace(hour=WORK_HOURS_START, minute=0)
    day_end    = datetime.strptime(date_str, "%Y-%m-%d").replace(hour=WORK_HOURS_END,   minute=0)

    while slot_start + timedelta(minutes=duration_minutes) <= day_end:
        slot_end = slot_start + timedelta(minutes=duration_minutes)
        conflict = any(
            slot_start < datetime.strptime(b["end"],   "%Y-%m-%d %H:%M") and
            slot_end   > datetime.strptime(b["start"], "%Y-%m-%d %H:%M")
            for b in busy
        )
        if not conflict:
            free.append(slot_start.strftime("%H:%M"))
        slot_start += timedelta(minutes=duration_minutes)

    return free


def get_free_slots_next_n_days(
    from_date: str,
    duration_minutes: int = 60,
    days: int = 5,
    max_slots_per_day: int = 4,
) -> dict[str, list[str]]:
    result    = {}
    current   = datetime.strptime(from_date, "%Y-%m-%d") + timedelta(days=1)
    days_done = 0

    while days_done < days:
        if current.weekday() < 5:
            date_str = current.strftime("%Y-%m-%d")
            slots    = get_all_free_slots(date_str, duration_minutes)[:max_slots_per_day]
            if slots:
                result[date_str] = slots
                days_done += 1
        current += timedelta(days=1)

    return result


def _reclassify_by_thread(email_type: str, thread_id: str) -> str:

    if email_type in ("acceptance", "rejection"):
        return email_type

    existing = _get_negotiation(thread_id)
    if existing and existing.get("status") == "active":
        if email_type == "fresh_request":
            print(f"[negotiation] Reclassified fresh_request → counter_proposal "
                  f"(active thread: {thread_id!r})")
            return "counter_proposal"

    return email_type

CLASSIFY_SYSTEM = """You are a meeting negotiation email classifier. Output JSON only. Never explain.

Classify the email into exactly one of these types:
- "fresh_request"    : first-time request to schedule a meeting (no prior thread)
- "counter_proposal" : cannot make the proposed time, offering alternative slots
- "acceptance"       : explicitly accepting a specific proposed time
- "rejection"        : declining entirely, ending the scheduling attempt
- "clarification"    : asking for more information before committing

Extract:
- proposed_slots  : list of {"date": "YYYY-MM-DD", "time": "HH:MM"} explicitly stated
- accepted_slot   : {"date": "YYYY-MM-DD", "time": "HH:MM"} if type is acceptance and
                    a specific slot is named, otherwise null
- proposed_date   : "YYYY-MM-DD" if a day is named without specific times, else ""
- duration_minutes: integer, default 60
- title           : meeting title or topic
- is_agent        : true if email appears automated (structured slot bullet lists,
                    agent signature tags, or machine-like formatting)

Output:
{
  "type": "...",
  "proposed_slots":  [{"date": "...", "time": "..."}],
  "accepted_slot":   {"date": "...", "time": "..."} or null,
  "proposed_date":   "YYYY-MM-DD or empty",
  "duration_minutes": 60,
  "title": "...",
  "is_agent": false
}"""


def classify_negotiation_email(email: dict) -> dict:
    today    = datetime.now().strftime("%A %d %B %Y")
    is_agent = is_agent_email(email.get("body", ""))

    user_msg = (
        f"Today is {today}.\n\n"
        f"From: {email['from']} <{email['email']}>\n"
        f"Subject: {email['subject']}\n"
        f"Body:\n{email['body']}\n\n"
        f"Classify this email. Output JSON only."
    )

    response = ask_llm_chat(CLASSIFY_SYSTEM, user_msg)
    result   = safe_parse_json(response)

    if not result:
        result = {}

    result.setdefault("type",             "fresh_request")
    result.setdefault("proposed_slots",   [])
    result.setdefault("accepted_slot",    None)
    result.setdefault("proposed_date",    "")
    result.setdefault("duration_minutes", 60)
    result.setdefault("title",            "Meeting")
    result["is_agent"] = is_agent or result.get("is_agent", False)

    result["proposed_slots"] = [
        s for s in result["proposed_slots"]
        if s.get("date") and s.get("time")
    ]

    ac = result.get("accepted_slot")
    if not isinstance(ac, dict) or not ac.get("date") or not ac.get("time"):
        result["accepted_slot"] = None

    return result

NEGOTIATION_REPLY_SYSTEM = """You are a professional scheduling assistant. Output JSON only.
Write a concise, professional reply email for the given negotiation scenario.
Output: {"subject": "...", "body": "..."}
Keep it brief and direct. Do not pad with unnecessary pleasantries."""


def _compose_negotiation_reply(
    original_email: dict,
    scenario: str,
    extra_context: str,
    is_agent_to_agent: bool,
) -> str:
    style = (
        "Machine-to-machine message. Be extremely concise and structured. "
        "List all time slots as bullet points in format:  • YYYY-MM-DD at HH:MM"
        if is_agent_to_agent else
        "Human recipient. Be professional and warm but brief."
    )

    user_msg = (
        f"Scenario: {scenario}\n"
        f"Style: {style}\n"
        f"Extra context: {extra_context or 'None'}\n"
        f"Replying to: {original_email['from']}\n"
        f"Original subject: {original_email['subject']}\n\n"
        f"Write the reply. Output JSON only."
    )

    response = ask_llm_chat(NEGOTIATION_REPLY_SYSTEM, user_msg)
    result   = safe_parse_json(response)

    if result and result.get("body"):
        return result["body"]

    return f"Regarding your meeting request — {scenario.lower()}"

def handle_fresh_request(email: dict, classification: dict, thread_id: str) -> dict:
    title    = classification.get("title", "Meeting")
    duration = int(classification.get("duration_minutes") or 60)
    slots    = classification.get("proposed_slots", [])
    is_agent = classification.get("is_agent", False)

    negotiation = {
        "thread_id":  thread_id,
        "title":      title,
        "duration":   duration,
        "initiator":  email["email"],
        "rounds":     0,
        "status":     "active",
        "history":    [],
        "created_at": datetime.now().isoformat(),
    }

    for slot in slots:
        date, time_str = slot.get("date", ""), slot.get("time", "")
        if not date or not time_str:
            continue

        avail = check_availability(date, time_str, duration)
        if avail["available"]:
            cal_result = calendar_agent(
                f"{title} with {email['from']} on {date} at {time_str} for {duration} minutes"
            )
            scenario = (
                f"Confirm that '{title}' on {date} at {time_str} for {duration} minutes "
                f"has been successfully booked."
            )
            body = _compose_negotiation_reply(email, scenario, "", is_agent)
            body = inject_agent_tag(body)

            negotiation["rounds"] = 1
            negotiation["history"].append({"round": 1, "action": "confirmed", "slot": slot})
            _upsert_negotiation(thread_id, negotiation)
            _close_negotiation(thread_id, "confirmed")

            return {
                "action":    "confirmed",
                "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
                "slot":      slot,
                "calendar":  cal_result,
                "thread_id": thread_id,
            }

    return _build_counter_proposal(email, classification, negotiation, thread_id, is_first=True)

def handle_counter_proposal(email: dict, classification: dict, thread_id: str) -> dict:
    negotiation = _get_negotiation(thread_id)
    if not negotiation:
        return handle_fresh_request(email, classification, thread_id)

    negotiation["rounds"] = negotiation.get("rounds", 0) + 1 # Round counter fix: increment here, once, before anything else
    is_agent = classification.get("is_agent", False)

    if negotiation["rounds"] > MAX_ROUNDS:
        return _handle_negotiation_failure(email, negotiation, thread_id, reason="max_rounds")

    slots    = classification.get("proposed_slots", [])
    duration = negotiation.get("duration", 60)

    for slot in slots:
        date, time_str = slot.get("date", ""), slot.get("time", "")
        if not date or not time_str:
            continue

        avail = check_availability(date, time_str, duration)
        if avail["available"]:
            cal_result = calendar_agent(
                f"{negotiation['title']} with {email['from']} on {date} at {time_str} for {duration} minutes"
            )
            scenario = (
                f"Confirm that '{negotiation['title']}' on {date} at {time_str} "
                f"has been booked successfully."
            )
            body = _compose_negotiation_reply(email, scenario, "", is_agent)
            body = inject_agent_tag(body)

            negotiation["history"].append({
                "round": negotiation["rounds"], "action": "confirmed", "slot": slot
            })
            _upsert_negotiation(thread_id, negotiation)
            _close_negotiation(thread_id, "confirmed")

            return {
                "action":    "confirmed",
                "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
                "slot":      slot,
                "calendar":  cal_result,
                "thread_id": thread_id,
            }

    # None worked — counter again. Pass negotiation with already-incremented round.
    return _build_counter_proposal(email, classification, negotiation, thread_id, is_first=False)


def _build_counter_proposal(
    email: dict,
    classification: dict,
    negotiation: dict,
    thread_id: str,
    is_first: bool,
) -> dict:
    duration = negotiation.get("duration", int(classification.get("duration_minutes") or 60))
    is_agent = classification.get("is_agent", False)

    if is_first:
        negotiation["rounds"] = negotiation.get("rounds", 0) + 1

    proposed_date = ""
    for s in classification.get("proposed_slots", []):
        if s.get("date"):
            proposed_date = s["date"]
            break
    if not proposed_date:
        proposed_date = classification.get("proposed_date", "")
    if not proposed_date:
        proposed_date = datetime.now().strftime("%Y-%m-%d")

    free_today = get_all_free_slots(proposed_date, duration)

    if free_today:
        slots_text    = "\n".join(f"  • {proposed_date} at {t}" for t in free_today)
        scenario      = (
            f"Cannot accommodate the requested time for '{negotiation.get('title','Meeting')}'. "
            f"Offer all free slots on {proposed_date}:\n{slots_text}"
        )
        counter_slots = [{"date": proposed_date, "time": t} for t in free_today]
    else:
        future_slots = get_free_slots_next_n_days(proposed_date, duration)
        if not future_slots:
            return _handle_negotiation_failure(email, negotiation, thread_id, reason="no_availability")

        lines, counter_slots = [], []
        for date, times in future_slots.items():
            for t in times:
                lines.append(f"  • {date} at {t}")
                counter_slots.append({"date": date, "time": t})

        slots_text = "\n".join(lines)
        scenario   = (
            f"No slots available on {proposed_date} for '{negotiation.get('title','Meeting')}'. "
            f"Offering alternative slots across the next available working days:\n{slots_text}"
        )

    body = _compose_negotiation_reply(email, scenario, "", is_agent)
    body = inject_agent_tag(body)

    negotiation["history"] = negotiation.get("history", [])
    negotiation["history"].append({
        "round":         negotiation["rounds"],
        "action":        "counter_proposal",
        "slots_offered": counter_slots,
    })
    negotiation["status"] = "active"
    _upsert_negotiation(thread_id, negotiation)

    return {
        "action":        "counter_proposal",
        "reply":         {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "slots_offered": counter_slots,
        "thread_id":     thread_id,
    }


def handle_acceptance(email: dict, thread_id: str) -> dict:
    """
    Try to book the slot the other party explicitly named first.
    Fall back to the slots from our last counter_proposal only if
    no explicit slot was stated.
    """
    negotiation    = _get_negotiation(thread_id)
    is_agent       = is_agent_email(email.get("body", ""))
    classification = classify_negotiation_email(email)
    duration       = (negotiation or {}).get("duration", 60)
    title          = (negotiation or {}).get("title", "Meeting")

    explicit   = classification.get("accepted_slot")
    candidates = []

    if explicit and explicit.get("date") and explicit.get("time"):
        candidates.append(explicit)

    for entry in reversed((negotiation or {}).get("history", [])):
        if entry.get("action") == "counter_proposal":
            candidates.extend(entry.get("slots_offered", []))
            break

    booked_slot = None
    cal_result  = None

    for slot in candidates:
        avail = check_availability(slot["date"], slot["time"], duration)
        if avail["available"]:
            cal_result  = calendar_agent(
                f"{title} with {email['from']} on {slot['date']} at {slot['time']} for {duration} minutes"
            )
            booked_slot = slot
            break

    scenario = (
        f"Confirm acceptance and booking of '{title}' on {booked_slot['date']} at {booked_slot['time']}."
        if booked_slot else
        f"The accepted slot is no longer available for '{title}'. Apologise and suggest manual rescheduling."
    )

    body = _compose_negotiation_reply(email, scenario, "", is_agent)
    body = inject_agent_tag(body)
    _close_negotiation(thread_id, "confirmed" if booked_slot else "failed")

    return {
        "action":    "confirmed" if booked_slot else "failed",
        "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "slot":      booked_slot,
        "calendar":  cal_result,
        "thread_id": thread_id,
    }


def handle_rejection(email: dict, thread_id: str) -> dict:
    is_agent = is_agent_email(email.get("body", ""))
    body     = _compose_negotiation_reply(
        email,
        "Acknowledge that the meeting will not proceed and close the thread politely.",
        "", is_agent
    )
    body = inject_agent_tag(body)
    _close_negotiation(thread_id, "rejected")

    return {
        "action":    "rejected",
        "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "thread_id": thread_id,
    }


def _handle_negotiation_failure(
    email: dict, negotiation: dict, thread_id: str, reason: str
) -> dict:
    is_agent  = is_agent_email(email.get("body", ""))
    scenarios = {
        "max_rounds":      "After several rounds no mutually agreeable time was found. "
                           "Suggest they reach out directly to resolve manually.",
        "no_availability": "No available slots exist in the coming days. "
                           "Request they propose a different week.",
    }
    body = _compose_negotiation_reply(
        email, scenarios.get(reason, "Negotiation could not be completed."), "", is_agent
    )
    body = inject_agent_tag(body)

    negotiation["history"] = negotiation.get("history", [])
    negotiation["history"].append({"action": "failed", "reason": reason})
    _upsert_negotiation(thread_id, negotiation)
    _close_negotiation(thread_id, "failed")

    return {
        "action":    "failed",
        "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "reason":    reason,
        "thread_id": thread_id,
        "escalate":  True,
    }

def negotiate(email: dict) -> dict | None:
    
    thread_id      = _extract_thread_id(email.get("subject", ""))
    classification = classify_negotiation_email(email)
    email_type     = _reclassify_by_thread(classification["type"], thread_id)
    classification["type"] = email_type

    current_round = (_get_negotiation(thread_id) or {}).get("rounds", 0) + 1
    print(
        f"[negotiation] Thread: {thread_id!r} | Type: {email_type} | "
        f"Agent: {classification['is_agent']} | Round: {current_round}"
    )

    if email_type == "acceptance":
        return handle_acceptance(email, thread_id)

    if email_type == "rejection":
        return handle_rejection(email, thread_id)

    existing = _get_negotiation(thread_id)

    if email_type == "fresh_request" and not existing:
        return handle_fresh_request(email, classification, thread_id)

    if email_type in ("counter_proposal", "fresh_request"):
        return handle_counter_proposal(email, classification, thread_id)

    print(f"[negotiation] Unhandled type '{email_type}' — requires manual review.")
    return None

def simulate_negotiation(
    title: str = "Project Sync",
    duration: int = 60,
    agent_a_name: str  = "Agent A",
    agent_a_email: str = "agent_a@example.com",
    agent_b_name: str  = "Agent B",
    agent_b_email: str = "agent_b@example.com",
    initial_slots: list[dict] | None = None,
    max_sim_rounds: int = 8,
) -> list[dict]:
    """
    Simulate a full negotiation between two agent instances entirely in-process.
    No real emails are sent. Uses the actual negotiate() logic on both sides,
    with in-memory email dicts passed back and forth.

    Returns a full transcript list where each entry is:
        {
            round:    int,
            sender:   str,
            receiver: str,
            action:   str,       # confirmed | counter_proposal | rejected | failed
            slot:     dict|None, # booked slot if confirmed
            email:    dict,      # the email that was processed this round
            result:   dict,      # full negotiate() return value
        }

    Usage:
        transcript = simulate_negotiation(
            title="Budget Review",
            initial_slots=[{"date": "2026-05-21", "time": "14:00"}]
        )
        for step in transcript:
            print(f"Round {step['round']} | {step['sender']} → {step['receiver']} | {step['action']}")
    """
    if not initial_slots:
        tomorrow      = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        initial_slots = [{"date": tomorrow, "time": "09:00"}]

    slots_text = "\n".join(f"  • {s['date']} at {s['time']}" for s in initial_slots)

    # Agent A sends the opening request targeting Agent B
    current_email = {
        "from":          agent_a_name,
        "email":         agent_a_email,
        "subject":       f"Meeting Request: {title}",
        "body":          (
            f"Hi,\n\nI'd like to schedule '{title}' ({duration} minutes).\n\n"
            f"I'm available at the following times:\n{slots_text}\n\n"
            f"Please confirm if any of these work.\n\nBest,\n{agent_a_name}\n\n"
            f"--\n{AGENT_TAG}"
        ),
        "received_time": datetime.now().isoformat(),
    }

    transcript      = []
    current_sender  = agent_a_name
    current_receiver = agent_b_name

    for sim_round in range(1, max_sim_rounds + 1):
        print(f"\n[sim] ── Round {sim_round}: {current_receiver} processing "
              f"email from {current_sender}")
        print(f"[sim]    Subject: {current_email['subject']}")

        result = negotiate(current_email)
        action = result["action"] if result else "manual_review"

        step = {
            "round":    sim_round,
            "sender":   current_sender,
            "receiver": current_receiver,
            "action":   action,
            "slot":     result.get("slot") if result else None,
            "email":    current_email,
            "result":   result,
        }
        transcript.append(step)

        if not result or action in ("confirmed", "rejected", "failed"):
            print(f"\n[sim] ══ Negotiation ended: {action} after {sim_round} round(s)")
            if result and result.get("slot"):
                s = result["slot"]
                print(f"[sim]    Booked slot: {s['date']} at {s['time']}")
            break

        # Swap sides: the reply becomes the next email
        reply = result["reply"]
        next_email = {
            "from":          current_receiver,
            "email":         agent_b_email if current_receiver == agent_b_name else agent_a_email,
            "subject":       reply["subject"],
            "body":          reply["body"],
            "received_time": datetime.now().isoformat(),
        }
        current_sender, current_receiver = current_receiver, current_sender
        current_email = next_email

    else:
        print(f"\n[sim] ══ Max sim rounds ({max_sim_rounds}) reached without resolution.")

    return transcript