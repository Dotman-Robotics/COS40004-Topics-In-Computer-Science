"""
negotiation_agent.py  -  Week 10 complete implementation + human negotiation extension

Key fixes in this version:
  1. Slots are embedded as a structured machine-readable block in every
     counter-proposal email - the LLM never paraphrases them. The receiving
     agent parses this block directly, eliminating the "1pm" / vanishing-slots bug.
  2. handle_acceptance validates that the accepted slot is actually free on
     THIS system's calendar before booking - prevents double-booking when the
     other system confirms a slot that is busy here.
  3. Round counter increments correctly in all paths.
  4. Email reclassification prevents fresh_request misrouting on reply threads.
  5. Simulation mode for offline testing.

Human negotiation extension (search HUMAN NEGOTIATION to find all related code):
  - Fully independent functions - zero changes to the agent-to-agent path above.
  - Entry point: negotiate_with_human(email) - drop-in parallel to negotiate(email).
  - Assumes human provides exact date/time in prose (e.g. "09:00 2026-07-06").
  - No slot blocks sent or expected. Replies use natural prose only.
  - Convergence via LLM accepted_slot extraction + offer history match (no Rule 0).
  - Sends ICS attachment string on confirmation so the human can import the event.
"""

import json
import os
import re
import threading
from datetime import datetime, timedelta

from ollama_client import ask_llm_chat
from calendar_agent import (
    check_availability, calendar_agent, get_busy_slots,
    book_tentative, confirm_tentative, cancel_tentative,
    _apply_defaults, _parse_event,
)
from utils import safe_parse_json

# -- Config --------------------------------------------------------------------

# Per-account state file so two instances in the same folder don't share state.
# e.g. negotiations_zoomertron_outlook_com.json
# Resolved at call time (not import time) to ensure --account CLI arg is always respected.
def _get_state_file() -> str:
    slug = os.environ.get("AGENT_ACCOUNT", "default").replace("@", "_").replace(".", "_")
    return f"negotiations_{slug}.json"

MAX_ROUNDS             = 5
MAX_COUNTER_SLOTS      = 3   # Max slots offered in any single counter-proposal email
AGENT_TAG              = "X-AgentSystem: true"

# Delimiter for the machine-readable slot block embedded in counter emails.
SLOT_BLOCK_START = "<<SLOTS_BEGIN>>"
SLOT_BLOCK_END   = "<<SLOTS_END>>"

WORK_HOURS_START = 8
WORK_HOURS_END   = 17  # Slots must START before 17:00 - matches calendar_agent.BUSINESS_HOURS_END

_state_lock = threading.Lock()

# How many days to keep closed (confirmed/rejected/failed) threads before
# pruning them from the state file. Active threads are always purged on startup.
CLOSED_THREAD_RETENTION_DAYS = 7


# -- State persistence ---------------------------------------------------------

def _load_state() -> dict:
    try:
        with open(_get_state_file(), "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_state(state: dict) -> None:
    with open(_get_state_file(), "w", encoding="utf-8") as f:
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
    # Prune old closed threads every time one closes
    _prune_closed_threads()


def get_all_negotiations() -> dict:
    with _state_lock:
        return _load_state()


# -- State cleanup -------------------------------------------------------------

def purge_active_negotiations() -> int:
    """
    Remove all threads with status 'active' from the state file.
    Called on startup to clear negotiations that never resolved
    (crashed mid-run, test runs, Outlook send failures, etc.).

    Also cancels any orphaned tentative calendar entries stored in those threads
    so they do not linger in Outlook.

    Returns the number of threads purged.
    """
    from calendar_agent import cancel_tentative

    with _state_lock:
        state   = _load_state()
        to_purge = {
            tid: data for tid, data in state.items()
            if data.get("status") == "active"
        }

        for tid, data in to_purge.items():
            entry_id = data.get("entry_id")
            title    = data.get("title", "Meeting")
            if entry_id:
                print(f"[negotiation] Cancelling orphaned tentative: '{title}' (thread: {tid!r})")
                try:
                    cancel_tentative(entry_id, title)
                except Exception as e:
                    print(f"[negotiation] Could not cancel tentative for {tid!r}: {e}")
            del state[tid]

        if to_purge:
            _save_state(state)
            print(f"[negotiation] Purged {len(to_purge)} stale active thread(s) on startup.")
        else:
            print(f"[negotiation] No stale active threads found.")

    return len(to_purge)


def _prune_closed_threads() -> None:
    """
    Remove confirmed/rejected/failed threads older than CLOSED_THREAD_RETENTION_DAYS.
    Called automatically every time a negotiation closes.
    Silent - does not print unless something is pruned.
    """
    cutoff = datetime.now().timestamp() - (CLOSED_THREAD_RETENTION_DAYS * 86400)

    with _state_lock:
        state    = _load_state()
        terminal = {"confirmed", "rejected", "failed"}
        to_prune = []

        for tid, data in state.items():
            if data.get("status") not in terminal:
                continue
            closed_at = data.get("closed_at", "")
            if not closed_at:
                continue
            try:
                closed_ts = datetime.fromisoformat(closed_at).timestamp()
                if closed_ts < cutoff:
                    to_prune.append(tid)
            except ValueError:
                continue

        for tid in to_prune:
            del state[tid]

        if to_prune:
            _save_state(state)
            print(f"[negotiation] Pruned {len(to_prune)} old closed thread(s) "
                  f"(>{CLOSED_THREAD_RETENTION_DAYS} days).")


# -- Thread ID -----------------------------------------------------------------

def _extract_thread_id(subject: str) -> str:
    """
    Derive a stable thread ID from the email subject.
    Strips ALL leading Re:/Fwd: prefixes so a deeply nested reply
    always maps to the same negotiation record as the original request.
    e.g. 'Re: Re: Re: Meeting Request: Project Sync'
      ->  'meeting_request:_project_sync'
    """
    while True:
        stripped = re.sub(r"^(Re|Fwd|FWD|RE|FW):\s*", "", subject, flags=re.IGNORECASE).strip()
        if stripped == subject:
            break
        subject = stripped
    return re.sub(r"\s+", "_", subject.lower())[:80]


# -- Agent detection -----------------------------------------------------------

def is_agent_email(body: str) -> bool:
    return AGENT_TAG.lower() in body.lower()


def inject_agent_tag(body: str) -> str:
    return body + f"\n\n--\n{AGENT_TAG}"


# -- Structured slot block -----------------------------------------------------
# Instead of asking the LLM to reproduce slot lists (which it paraphrases),
# we embed a machine-readable JSON block at the bottom of every counter-proposal.
# The receiving agent extracts this block directly - no LLM involved.

def _embed_slot_block(body: str, slots: list[dict], duration: int) -> str:
    """Append a structured slot block to an email body."""
    block = json.dumps({"slots": slots, "duration_minutes": duration})
    return f"{body}\n\n{SLOT_BLOCK_START}\n{block}\n{SLOT_BLOCK_END}"


def _extract_slot_block(body: str) -> list[dict]:
    """
    Extract slots from the structured block in an email body.
    Handles both multi-line (agent-generated) and single-line
    (Outlook-collapsed) formats.
    Returns an empty list if no block is present (e.g. from human senders).
    """
    # Outlook sometimes collapses newlines - match with optional whitespace
    match = re.search(
        re.escape(SLOT_BLOCK_START) + r"\s*(.*?)\s*" + re.escape(SLOT_BLOCK_END),
        body, re.DOTALL
    )
    if not match:
        return []
    raw = match.group(1).strip()
    try:
        data = json.loads(raw)
        return [s for s in data.get("slots", []) if s.get("date") and s.get("time")]
    except (json.JSONDecodeError, AttributeError):
        return []


def _extract_duration_from_block(body: str) -> int | None:
    """Extract duration_minutes from the slot block if present."""
    match = re.search(
        re.escape(SLOT_BLOCK_START) + r"\s*(.*?)\s*" + re.escape(SLOT_BLOCK_END),
        body, re.DOTALL
    )
    if not match:
        return None
    raw = match.group(1).strip()
    try:
        data = json.loads(raw)
        return int(data.get("duration_minutes", 60))
    except (json.JSONDecodeError, ValueError, AttributeError):
        return None


# -- Availability helpers ------------------------------------------------------

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
    max_slots_per_day: int = 2, #Note this is where the daily slots are changed
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


# -- Email reclassification ----------------------------------------------------

def _reclassify_by_thread(email_type: str, thread_id: str, body: str = "") -> str:
    """
    Reclassify based on thread state to prevent misrouting.

    Rule 0 - MUTUAL TENTATIVE (highest priority)
      If we have a tentative booking and the incoming email contains that
      exact same slot, both parties agree - this is an acceptance regardless
      of what the LLM classified it as. This is the definitive signal that
      negotiations have converged.

    Rule 1 - fresh_request on active thread -> counter_proposal

    Rule 2 - acceptance with slot block:
      - Slots match what we last offered -> acceptance
      - Slots are new -> counter_proposal
      - No history -> trust the classifier

    Rule 3 - rejection is always terminal
    """
    existing = _get_negotiation(thread_id)
    active   = existing and existing.get("status") == "active"

    if email_type == "rejection":
        return "rejection"

    if active:
        block_slots = _extract_slot_block(body)

        # Rule 0 - check for mutual tentative convergence
        # Find our own tentative slot from history
        our_tentative = None
        for entry in reversed(existing.get("history", [])):
            if entry.get("action") == "tentative" and entry.get("slot"):
                our_tentative = entry["slot"]
                break

        if our_tentative and block_slots:
            incoming = {(s["date"], s["time"]) for s in block_slots}
            our_key  = (our_tentative["date"], our_tentative["time"])
            if our_key in incoming:
                print(f"[negotiation] Mutual tentative detected at "
                      f"{our_tentative['date']} {our_tentative['time']} -> acceptance")
                return "acceptance"

        # Rule 2 - acceptance with slot block comparison
        if email_type == "acceptance" and block_slots:
            last_offered = set()
            for entry in reversed(existing.get("history", [])):
                if entry.get("action") == "tentative" and entry.get("slot"):
                    s = entry["slot"]
                    last_offered.add((s["date"], s["time"]))
                    break
                if entry.get("action") == "counter_proposal":
                    for s in entry.get("slots_offered", []):
                        last_offered.add((s["date"], s["time"]))
                    break

            if not last_offered:
                print(f"[negotiation] No offer history - trusting classifier: acceptance")
                return "acceptance"

            incoming = {(s["date"], s["time"]) for s in block_slots}
            if incoming.issubset(last_offered):
                print(f"[negotiation] Slot block matches our offer -> acceptance")
                return "acceptance"
            else:
                print(f"[negotiation] Slot block contains new slots -> counter_proposal")
                return "counter_proposal"

        if email_type == "acceptance" and not block_slots:
            return "acceptance"

        # Rule 1
        if email_type == "fresh_request":
            print(f"[negotiation] Reclassified fresh_request -> counter_proposal "
                  f"(active thread: {thread_id!r})")
            return "counter_proposal"

    return email_type


# -- Email classification ------------------------------------------------------

CLASSIFY_SYSTEM = """You are a meeting negotiation email classifier. Output JSON only. Never explain.

Classify the email into exactly one of these types:
- "fresh_request"    : first-time request to schedule a meeting
- "counter_proposal" : cannot make the time, offering alternatives
- "acceptance"       : explicitly accepting a proposed time
- "rejection"        : declining entirely
- "clarification"    : asking for more info

Extract:
- proposed_slots  : list of {"date": "YYYY-MM-DD", "time": "HH:MM"} explicitly stated
- accepted_slot   : {"date": "YYYY-MM-DD", "time": "HH:MM"} if acceptance and slot is named, else null
- proposed_date   : "YYYY-MM-DD" if a day is named without times, else ""
- duration_minutes: integer, default 60
- title           : meeting title
- is_agent        : true if email looks automated (structured slot lists, agent tags)

Output:
{
  "type": "...",
  "proposed_slots":  [{"date": "...", "time": "..."}],
  "accepted_slot":   {"date": "...", "time": "..."} or null,
  "proposed_date":   "",
  "duration_minutes": 60,
  "title": "...",
  "is_agent": false
}"""


def classify_negotiation_email(email: dict) -> dict:
    today    = datetime.now().strftime("%A %d %B %Y")
    is_agent = is_agent_email(email.get("body", ""))

    # -- Pull slots directly from the machine-readable block first -------------
    # This bypasses the LLM entirely for agent-to-agent emails, which is both
    # faster and eliminates misclassification from paraphrased slot descriptions.
    block_slots = _extract_slot_block(email.get("body", ""))
    block_duration = _extract_duration_from_block(email.get("body", ""))

    user_msg = (
        f"Today is {today}.\n\n"
        f"From: {email['from']} <{email['email']}>\n"
        f"Subject: {email['subject']}\n"
        f"Body:\n{email['body']}\n\n"
        f"Classify this email. Output JSON only."
    )

    response = ask_llm_chat(CLASSIFY_SYSTEM, user_msg)
    result   = safe_parse_json(response) or {}

    result.setdefault("type",             "fresh_request")
    result.setdefault("proposed_slots",   [])
    result.setdefault("accepted_slot",    None)
    result.setdefault("proposed_date",    "")
    result.setdefault("duration_minutes", 60)
    result.setdefault("title",            "Meeting")
    result["is_agent"] = is_agent or result.get("is_agent", False)

    # If the email contained a structured slot block, override LLM-extracted slots
    # with the authoritative machine-readable version
    if block_slots:
        result["proposed_slots"] = block_slots
        print(f"[negotiation] Used slot block: {len(block_slots)} slot(s) extracted directly")
    else:
        result["proposed_slots"] = [
            s for s in result["proposed_slots"]
            if s.get("date") and s.get("time")
        ]

    if block_duration:
        result["duration_minutes"] = block_duration

    ac = result.get("accepted_slot")
    if not isinstance(ac, dict) or not ac.get("date") or not ac.get("time"):
        result["accepted_slot"] = None

    return result


# -- Reply composers -----------------------------------------------------------
# Two separate prompts: one for counter-proposals (proposing language only),
# one for confirmations (confirming language only).
# Keeping them separate prevents the LLM from mixing tone and causing
# System A to misread a proposal as a confirmed booking.

COUNTER_PROPOSAL_REPLY_SYSTEM = """You are a professional scheduling assistant. Output JSON only.
Write a brief reply informing the other party that the requested time is unavailable
and that you are proposing alternative times for their consideration.

CRITICAL RULES:
- Use proposing language ONLY: "I'd like to suggest an alternative time", "Would any of the following options suit you?", "Please let me know if one of these works for you"
- NEVER use confirming language: never say "confirmed", "booked", "scheduled", "I've arranged"
- NEVER mention any specific times, dates, or slot data - not even as examples. The slots are listed separately below the message body.
- Do NOT include JSON, curly braces, or structured data in the body text
- Keep it to 2 sentences maximum

BAD example (do NOT do this): "I'd like to suggest three slots: {time_slot_1: 10:00 AM, time_slot_2: 11:30 AM}"
GOOD example: "I'd like to suggest an alternative time slot that may work for you. Would any of the following options suit you? Please let me know if one of these works for you."

Output: {"subject": "Re: <original subject>", "body": "..."}"""

CONFIRMATION_REPLY_SYSTEM = """You are a professional scheduling assistant. Output JSON only.
Write a brief reply confirming that a meeting has been successfully booked.

CRITICAL RULES:
- Use confirming language: "confirmed", "booked", "scheduled"
- Keep it to 1-2 sentences maximum
- Do NOT mention specific times - those are handled separately

Output: {"subject": "Re: <original subject>", "body": "..."}"""

REJECTION_REPLY_SYSTEM = """You are a professional scheduling assistant. Output JSON only.
Write a brief, polite reply acknowledging that the meeting cannot proceed.
Keep it to 1-2 sentences. Output: {"subject": "Re: <original subject>", "body": "..."}"""


def _compose_counter_reply(original_email: dict, title: str, n_slots: int) -> str:
    """Write a proposal reply - never a confirmation."""
    user_msg = (
        f"Original subject: {original_email['subject']}\n"
        f"Meeting title: {title}\n"
        f"Number of alternative slots being offered: {n_slots}\n"
        f"Write a brief reply proposing these alternatives. Output JSON only."
    )
    response = ask_llm_chat(COUNTER_PROPOSAL_REPLY_SYSTEM, user_msg)
    result   = safe_parse_json(response)
    if result and result.get("body"):
        return result["body"]
    return (
        f"Thank you for your message. Unfortunately the requested time isn't available. "
        f"I've listed {n_slots} alternative slot(s) below - please let me know if any work for you."
    )


def _compose_confirmation_reply(original_email: dict, title: str, date: str, time_str: str) -> str:
    """Write a confirmation reply - never a proposal."""
    user_msg = (
        f"Original subject: {original_email['subject']}\n"
        f"Meeting: {title} has been confirmed for {date} at {time_str}.\n"
        f"Write a brief confirmation. Output JSON only."
    )
    response = ask_llm_chat(CONFIRMATION_REPLY_SYSTEM, user_msg)
    result   = safe_parse_json(response)
    if result and result.get("body"):
        return result["body"]
    return f"Confirmed - '{title}' is booked for {date} at {time_str}. Looking forward to it."


def _compose_rejection_reply(original_email: dict, reason: str) -> str:
    """Write a rejection/failure reply."""
    user_msg = (
        f"Original subject: {original_email['subject']}\n"
        f"Reason: {reason}\n"
        f"Write a brief polite reply. Output JSON only."
    )
    response = ask_llm_chat(REJECTION_REPLY_SYSTEM, user_msg)
    result   = safe_parse_json(response)
    if result and result.get("body"):
        return result["body"]
    return "Unfortunately we were unable to find a mutually agreeable time. Please reach out to schedule manually."


# -- Core negotiation handlers -------------------------------------------------

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
            start_dt = datetime.strptime(f"{date} {time_str}", "%Y-%m-%d %H:%M")
            end_dt   = start_dt + timedelta(minutes=duration)
            event    = {"title": title, "duration_minutes": duration,
                        "description": f"Meeting with {email['from']}", "location": ""}
            tent     = book_tentative(event, start_dt, end_dt)

            body = _compose_counter_reply(email, title, 1)
            # For a fresh request where we CAN do the proposed time,
            # we still need THEIR confirmation - so this is a "yes, works for us"
            # counter, not a final booking. We embed the agreed slot so they can confirm.
            agreed_slot = {"date": date, "time": time_str}
            body = _embed_slot_block(body, [agreed_slot], duration)
            body = inject_agent_tag(body)

            negotiation["rounds"]   = 1
            negotiation["entry_id"] = tent.get("entry_id")
            negotiation["history"].append({
                "round": 1, "action": "tentative", "slot": agreed_slot,
                "entry_id": tent.get("entry_id"),
            })
            _upsert_negotiation(thread_id, negotiation)

            return {
                "action":    "tentative",
                "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
                "slot":      agreed_slot,
                "entry_id":  tent.get("entry_id"),
                "thread_id": thread_id,
            }

    return _build_counter_proposal(email, classification, negotiation, thread_id, is_first=True)


def handle_counter_proposal(email: dict, classification: dict, thread_id: str) -> dict:
    negotiation = _get_negotiation(thread_id)

    if not negotiation:
        # No record exists - this system is the initiator receiving a reply.
        # The initiator sent the opening email manually so never wrote a record.
        # Create one now so the round counter and state tracking work correctly.
        print(f"[negotiation] No record for {thread_id!r} - creating initiator record.")
        negotiation = {
            "thread_id":  thread_id,
            "title":      classification.get("title", "Meeting"),
            "duration":   int(classification.get("duration_minutes") or 60),
            "initiator":  email["email"],   # the other party is now the sender
            "rounds":     0,
            "status":     "active",
            "history":    [],
            "entry_id":   None,
            "created_at": datetime.now().isoformat(),
        }
        _upsert_negotiation(thread_id, negotiation)

    negotiation["rounds"] = negotiation.get("rounds", 0) + 1
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
            start_dt = datetime.strptime(f"{date} {time_str}", "%Y-%m-%d %H:%M")
            end_dt   = start_dt + timedelta(minutes=duration)
            event    = {"title": negotiation['title'], "duration_minutes": duration,
                        "description": f"Meeting with {email['from']}", "location": ""}
            tent     = book_tentative(event, start_dt, end_dt)

            # We can do this slot - propose it back and wait for their acceptance
            agreed_slot = {"date": date, "time": time_str}
            body = _compose_counter_reply(email, negotiation['title'], 1)
            body = _embed_slot_block(body, [agreed_slot], duration)
            body = inject_agent_tag(body)

            negotiation["entry_id"] = tent.get("entry_id")
            negotiation["history"].append({
                "round": negotiation["rounds"], "action": "tentative",
                "slot": agreed_slot, "entry_id": tent.get("entry_id"),
            })
            _upsert_negotiation(thread_id, negotiation)

            return {
                "action":    "tentative",
                "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
                "slot":      agreed_slot,
                "entry_id":  tent.get("entry_id"),
                "thread_id": thread_id,
            }

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

    # Determine reference date
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
        counter_slots = [{"date": proposed_date, "time": t} for t in free_today[:MAX_COUNTER_SLOTS]]
        scenario = (
            f"Cannot accommodate the requested time for '{negotiation.get('title','Meeting')}' "
            f"on {proposed_date}. Offering {len(counter_slots)} alternative slot(s)."
        )
    else:
        future_slots = get_free_slots_next_n_days(proposed_date, duration)
        if not future_slots:
            return _handle_negotiation_failure(email, negotiation, thread_id, reason="no_availability")

        counter_slots = []
        for date, times in future_slots.items():
            for t in times:
                counter_slots.append({"date": date, "time": t})
                if len(counter_slots) >= MAX_COUNTER_SLOTS:
                    break
            if len(counter_slots) >= MAX_COUNTER_SLOTS:
                break

        scenario = (
            f"No slots available on {proposed_date} for '{negotiation.get('title','Meeting')}'. "
            f"Offering {len(counter_slots)} slot(s) across the next available working days."
        )

    # -- Compose prose body (LLM) then embed authoritative slot block ----------
    # The LLM writes a brief message only - it is explicitly told NOT to list slots.
    # The structured block carries the actual slot data the receiving agent will parse.
    body = _compose_counter_reply(email, negotiation.get("title", "Meeting"), len(counter_slots))

    if is_agent:
        slot_lines = "\n".join(f"  • {s['date']} at {s['time']}" for s in counter_slots)
        body = body + f"\n\nAvailable slots:\n{slot_lines}"

    body = _embed_slot_block(body, counter_slots, duration)
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
    The other party has accepted our proposed slot.
    Promote the tentative calendar entry to confirmed, then reply.
    If the slot is now gone (race condition), re-open with fresh slots.
    """
    negotiation    = _get_negotiation(thread_id)
    classification = classify_negotiation_email(email)
    duration       = (negotiation or {}).get("duration", 60)
    title          = (negotiation or {}).get("title", "Meeting")
    entry_id       = (negotiation or {}).get("entry_id")

    # Find the tentative slot from history
    tentative_slot = None
    for entry in reversed((negotiation or {}).get("history", [])):
        if entry.get("action") == "tentative" and entry.get("slot"):
            tentative_slot = entry["slot"]
            break

    # Also check explicit accepted_slot from classifier as a fallback
    explicit = classification.get("accepted_slot")

    slot_to_confirm = tentative_slot or explicit

    if slot_to_confirm:
        # Verify it is still free (race condition guard)
        avail = check_availability(
            slot_to_confirm["date"], slot_to_confirm["time"], duration
        )
        # If tentative blocks are the only conflict, it IS still ours - still confirm
        conflicts = avail.get("conflicts", [])
        only_our_tentative = (
            not avail["available"] and
            all(c.get("tentative") and title.lower() in c.get("title", "").lower()
                for c in conflicts)
        )

        if avail["available"] or only_our_tentative:
            # Promote the tentative entry
            if entry_id:
                confirm_status = confirm_tentative(entry_id, title)
                print(f"[negotiation] {confirm_status}")
            else:
                # No entry_id - create a fresh confirmed booking
                from calendar_agent import calendar_agent as _cal
                _cal(f"{title} with {email['from']} on {slot_to_confirm['date']} "
                     f"at {slot_to_confirm['time']} for {duration} minutes")

            body = _compose_confirmation_reply(
                email, title, slot_to_confirm["date"], slot_to_confirm["time"]
            )
            body = inject_agent_tag(body)
            _close_negotiation(thread_id, "confirmed")

            return {
                "action":    "confirmed",
                "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
                "slot":      slot_to_confirm,
                "thread_id": thread_id,
            }

        else:
            # Slot is now taken by something else - cancel tentative and re-negotiate
            if entry_id:
                cancel_tentative(entry_id, title)
            print(f"[negotiation] Tentative slot now conflicts - re-opening negotiation.")

    # Re-open with fresh slots
    if negotiation:
        negotiation["status"]   = "active"
        negotiation["entry_id"] = None
        negotiation["rounds"]   = negotiation.get("rounds", 0) + 1
        _upsert_negotiation(thread_id, negotiation)

    ref_date = datetime.now().strftime("%Y-%m-%d")
    free     = get_all_free_slots(ref_date, duration)
    if not free:
        future      = get_free_slots_next_n_days(ref_date, duration)
        fresh_slots = [{"date": d, "time": t} for d, times in future.items() for t in times]
    else:
        fresh_slots = [{"date": ref_date, "time": t} for t in free]

    body = _compose_counter_reply(email, title, len(fresh_slots))
    if fresh_slots:
        body = _embed_slot_block(body, fresh_slots, duration)
    body = inject_agent_tag(body)

    return {
        "action":        "counter_proposal",
        "reply":         {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "slots_offered": fresh_slots,
        "slot":          None,
        "thread_id":     thread_id,
    }


def handle_rejection(email: dict, thread_id: str) -> dict:
    negotiation = _get_negotiation(thread_id)
    entry_id    = (negotiation or {}).get("entry_id")
    title       = (negotiation or {}).get("title", "Meeting")

    if entry_id:
        print(f"[negotiation] Rejection received - cancelling tentative booking.")
        cancel_tentative(entry_id, title)

    body = _compose_rejection_reply(email, "The other party has declined to proceed.")
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
    entry_id = negotiation.get("entry_id")
    title    = negotiation.get("title", "Meeting")

    if entry_id:
        print(f"[negotiation] Negotiation failed ({reason}) - cancelling tentative booking.")
        cancel_tentative(entry_id, title)

    reasons = {
        "max_rounds":      "After several rounds no mutually agreeable time was found. Please reach out directly.",
        "no_availability": "No available slots exist in the coming days. Please propose a different week.",
    }
    body = _compose_rejection_reply(email, reasons.get(reason, "Negotiation could not be completed."))
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


# -- Main entry point ----------------------------------------------------------

def negotiate(email: dict) -> dict | None:
    thread_id = _extract_thread_id(email.get("subject", ""))

    # Guard: ignore any email on an already-closed thread.
    # Prevents courtesy confirmation replies from triggering duplicate bookings.
    existing = _get_negotiation(thread_id)
    if existing and existing.get("status") in ("confirmed", "rejected", "failed"):
        print(f"[negotiation] Thread {thread_id!r} already closed "
              f"({existing['status']}) - ignoring email.")
        return None

    classification = classify_negotiation_email(email)
    email_type     = _reclassify_by_thread(classification["type"], thread_id, email.get("body", ""))
    classification["type"] = email_type

    current_round = (existing or {}).get("rounds", 0) + 1
    print(
        f"[negotiation] Thread: {thread_id!r} | Type: {email_type} | "
        f"Agent: {classification['is_agent']} | Round: {current_round} | "
        f"Slots in block: {len(_extract_slot_block(email.get('body','')))}"
    )

    if email_type == "acceptance":
        return handle_acceptance(email, thread_id)
    if email_type == "rejection":
        return handle_rejection(email, thread_id)

    if email_type == "fresh_request" and not existing:
        return handle_fresh_request(email, classification, thread_id)
    if email_type in ("counter_proposal", "fresh_request"):
        return handle_counter_proposal(email, classification, thread_id)

    print(f"[negotiation] Unhandled type '{email_type}' - requires manual review.")
    return None


# -- Simulation mode -----------------------------------------------------------

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
    No real emails sent. Uses live Outlook calendar for availability checks.
    """
    if not initial_slots:
        tomorrow      = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        initial_slots = [{"date": tomorrow, "time": "09:00"}]

    slots_text = "\n".join(f"  • {s['date']} at {s['time']}" for s in initial_slots)
    body = (
        f"Hi,\n\nI'd like to schedule '{title}' ({duration} minutes).\n\n"
        f"I'm available at the following times:\n{slots_text}\n\n"
        f"Please confirm if any of these work.\n\nBest,\n{agent_a_name}"
    )
    body = _embed_slot_block(body, initial_slots, duration)
    body = inject_agent_tag(body)

    current_email = {
        "from":          agent_a_name,
        "email":         agent_a_email,
        "subject":       f"Meeting Request: {title}",
        "body":          body,
        "received_time": datetime.now().isoformat(),
    }

    transcript       = []
    current_sender   = agent_a_name
    current_receiver = agent_b_name

    for sim_round in range(1, max_sim_rounds + 1):
        print(f"\n[sim] -- Round {sim_round}: {current_receiver} processing "
              f"email from {current_sender}")

        result = negotiate(current_email)
        action = result["action"] if result else "manual_review"

        transcript.append({
            "round":    sim_round,
            "sender":   current_sender,
            "receiver": current_receiver,
            "action":   action,
            "slot":     result.get("slot") if result else None,
            "email":    current_email,
            "result":   result,
        })

        if not result or action in ("confirmed", "rejected", "failed"):
            print(f"\n[sim] -- Ended: {action} after {sim_round} round(s)")
            if result and result.get("slot"):
                s = result["slot"]
                print(f"[sim]    Booked: {s['date']} at {s['time']}")
            break

        reply = result["reply"]
        current_email = {
            "from":          current_receiver,
            "email":         agent_b_email if current_receiver == agent_b_name else agent_a_email,
            "subject":       reply["subject"],
            "body":          reply["body"],
            "received_time": datetime.now().isoformat(),
        }
        current_sender, current_receiver = current_receiver, current_sender

    else:
        print(f"\n[sim] -- Max sim rounds ({max_sim_rounds}) reached without resolution.")

    return transcript

# +------------------------------------------------------------------------------+
# |  HUMAN NEGOTIATION - independent extension                                  |
# |                                                                              |
# |  All functions below are exclusively for human ↔ agent negotiation.         |
# |  None of them are called by, or call into, the agent-to-agent path above.   |
# |  The agent-to-agent path (negotiate(), handle_fresh_request(), etc.)        |
# |  is completely untouched.                                                    |
# |                                                                              |
# |  Entry point: negotiate_with_human(email)                                   |
# |  Called by email_monitor.py instead of negotiate() when is_agent == False.  |
# |                                                                              |
# |  Assumptions about the human sender:                                         |
# |  - Provides exact date and time in prose, e.g. "09:00 2026-07-06"           |
# |  - Does not send or expect slot blocks (<<SLOTS_BEGIN>> / <<SLOTS_END>>)    |
# |  - Does not send or expect X-AgentSystem: true tags                         |
# |  - Replies are natural language - "Yes, that works", "Can we do Friday?"    |
# |                                                                              |
# |  Convergence strategy (no Rule 0):                                          |
# |  LLM classifies reply as acceptance + accepted_slot matches last offer      |
# |  -> confirmed. No mutual tentative required.                                 |
# +------------------------------------------------------------------------------+


# -- HUMAN NEGOTIATION: LLM prompts -------------------------------------------
# Separate from agent prompts - these produce natural prose, never slot blocks.

# HUMAN NEGOTIATION: classifier prompt - identical intent to CLASSIFY_SYSTEM
# but explicitly tuned for human prose with casual date/time phrasing.
HUMAN_CLASSIFY_SYSTEM = """You are a meeting scheduling email classifier. Output JSON only. Never explain.

The sender is a HUMAN - not an automated system. They may use casual language.
Date/time format used by this human: time first then date, e.g. "09:00 2026-07-06".
Always parse times as 24-hour HH:MM format and dates as YYYY-MM-DD.

IMPORTANT: The email body contains ONLY the sender's newest message - quoted reply text has been stripped.
Extract dates and times from the sender's words only.

CRITICAL SLOT EXTRACTION RULES:
- Each slot must pair the correct time WITH the correct date from the same phrase
- "11:00 2026-07-06" means time=11:00 date=2026-07-06 - never swap them
- "08:00 2026-07-07" means time=08:00 date=2026-07-07 - never swap them
- If the message says "X doesn't work, can we do Y", extract BOTH X and Y as separate slots
- Do NOT mix the time from one slot with the date from another

BAD: "11:00 2026-07-06 doesn't work, can we do 08:00 2026-07-07"
     -> extracted as [06/08:00, 07/11:00]  WRONG - times and dates are swapped
GOOD: same input
     -> extracted as [06/11:00, 07/08:00]  CORRECT - each time stays with its date

Classify the email into exactly one of:
- "fresh_request"    : first-time request to schedule a meeting
- "counter_proposal" : cannot make the time, offering or requesting alternatives
- "acceptance"       : explicitly agreeing to a proposed time
- "rejection"        : declining to meet entirely
- "clarification"    : asking a question before committing

Extract:
- proposed_slots  : list of {"date": "YYYY-MM-DD", "time": "HH:MM"} - all slots mentioned, correctly paired
- accepted_slot   : {"date": "YYYY-MM-DD", "time": "HH:MM"} if this is an acceptance, else null
- proposed_date   : "YYYY-MM-DD" if a day is mentioned without a time, else ""
- duration_minutes: integer - default 60 if not stated
- title           : meeting title or topic

Output JSON:
{
  "type": "...",
  "proposed_slots":   [{"date": "...", "time": "..."}],
  "accepted_slot":    {"date": "...", "time": "..."} or null,
  "proposed_date":    "",
  "duration_minutes": 60,
  "title":            "..."
}"""


# HUMAN NEGOTIATION: counter-proposal prompt - writes natural prose with slots
# inline as readable text. No JSON blocks, no bullet machine-lists.
HUMAN_COUNTER_PROPOSAL_SYSTEM = """You are a professional scheduling assistant writing on behalf of a business.
Output JSON only. Never explain outside the JSON.

The recipient is a HUMAN. Write a warm, natural reply in plain prose.
The originally requested time is not available. However you HAVE found alternative times that ARE available.
Your job is to politely say the requested time doesn't work, then OFFER the alternative times as times YOU ARE FREE.

RULES:
- First sentence: briefly apologise that the requested time is unavailable
- Second sentence: say you ARE available at the given times and ask if any suit them
  e.g. "I'm available on 6 July at 11:00 AM or 12:00 PM, or 7 July at 8:00 AM - would any of these work for you?"
- Mention ONLY the exact times given to you - do not invent or add any other times
- Do NOT say the alternatives are unavailable - they ARE available
- Do NOT use bullet points, JSON, or structured formatting
- Keep it to 2 sentences maximum

BAD: "I'm afraid 6 July at 11:00 AM is not available" - the alternatives ARE available, never say they aren't
GOOD: "Unfortunately that time doesn't work, but I'm free on 6 July at 11:00 AM or 12:00 PM - would either suit you?"

Output: {"subject": "Re: <original subject>", "body": "..."}"""

# HUMAN NEGOTIATION: confirmation prompt - natural prose, includes full date/time.
HUMAN_CONFIRMATION_SYSTEM = """You are a professional scheduling assistant writing on behalf of a business.
Output JSON only. Never explain outside the JSON.

The recipient is a HUMAN. Write a warm confirmation that the meeting is booked.

RULES:
- State the meeting title, date, and time clearly in plain prose
  e.g. "Great - Project Sync is confirmed for Monday 6 July at 9:00 AM. Looking forward to it!"
- Use friendly professional language
- Mention that a calendar invite is attached
- Keep it to 2 sentences maximum

Output: {"subject": "Re: <original subject>", "body": "..."}"""


# HUMAN NEGOTIATION: rejection prompt - polite, no technical language.
HUMAN_REJECTION_SYSTEM = """You are a professional scheduling assistant writing on behalf of a business.
Output JSON only. Never explain outside the JSON.

The recipient is a HUMAN. Write a polite reply that the meeting cannot be scheduled.

RULES:
- Be warm and apologetic
- Suggest they reach out directly to reschedule manually
- No technical language, no slot blocks, no JSON in the body
- Keep it to 2 sentences maximum

Output: {"subject": "Re: <original subject>", "body": "..."}"""


# HUMAN NEGOTIATION: tentative booking reply prompt.
# Used when a slot is provisionally held - asks human to confirm, not pick from options.
HUMAN_TENTATIVE_SYSTEM = """You are a professional scheduling assistant writing on behalf of a business.
Output JSON only. Never explain outside the JSON.

The recipient is a HUMAN. A specific time slot has been provisionally held for the meeting.
Write a warm reply saying you have provisionally held that time and asking them to confirm.

RULES:
- State the exact date and time provided to you
- Say it has been provisionally held or reserved
- Ask them to confirm whether it works
- Do NOT offer alternatives - only the one provisionally held slot
- Friendly professional tone, 2 sentences maximum

GOOD: "I've provisionally held 7 July at 10:00 AM for our meeting - does that work for you?"
BAD: "I'm free on 7 July at 10:00 AM or 8 July at 11:30 AM"

Output: {"subject": "Re: <original subject>", "body": "..."}"""


def _human_compose_tentative_request(
    original_email: dict,
    title: str,
    slot: dict,
) -> str:
    """
    HUMAN NEGOTIATION: compose a reply asking the human to confirm a provisionally held slot.
    Distinct from _human_compose_counter which offers multiple options.
    """
    dt = datetime.strptime(f"{slot['date']} {slot['time']}", "%Y-%m-%d %H:%M")
    slot_phrase = _fmt_slot_human(dt)
    user_msg = (
        f"Original subject: {original_email['subject']}\n"
        f"Meeting title: {title}\n"
        f"Provisionally held slot: {slot_phrase}\n"
        f"Write a reply asking the human to confirm this slot. Output JSON only."
    )
    response = ask_llm_chat(HUMAN_TENTATIVE_SYSTEM, user_msg)
    result   = safe_parse_json(response)
    if result and result.get("body"):
        body = result["body"]
        # Verify the slot phrase appears in the body
        if slot_phrase.split(" at ")[1] in body:
            return body
    # Fallback - deterministic prose
    return (
        f"I've provisionally held {slot_phrase} for {title} - "
        f"does that work for you?"
    )




def _fmt_slot_human(dt) -> str:
    """
    HUMAN NEGOTIATION: format a datetime as readable prose for a human recipient.
    Uses portable strftime codes - Windows does not support %-d or %-I.
    e.g. datetime(2026, 7, 6, 8, 0) -> '6 July at 8:00 AM'
    """
    hour_str = dt.strftime("%I").lstrip("0") or "12"
    return f"{dt.day} {dt.strftime('%B')} at {hour_str}:{dt.strftime('%M %p')}"


def _fmt_slot_human_long(dt) -> str:
    """
    HUMAN NEGOTIATION: format a datetime with full weekday and year.
    e.g. datetime(2026, 7, 6, 8, 0) -> 'Monday 6 July 2026 at 8:00 AM'
    """
    hour_str = dt.strftime("%I").lstrip("0") or "12"
    return f"{dt.strftime('%A')} {dt.day} {dt.strftime('%B %Y')} at {hour_str}:{dt.strftime('%M %p')}"


# -- HUMAN NEGOTIATION: email composers ---------------------------------------

def _human_classify_email(email: dict) -> dict:
    """
    HUMAN NEGOTIATION: classify an email from a human sender.
    Slot extraction is done via regex (deterministic) — the LLM is only
    used for intent classification (fresh_request / acceptance / etc.).
    This prevents hallucinated dates from Llama 3.2.
    """
    import re as _re

    # Extract only the sender's new text (first paragraph, CRLF-aware)
    raw_body = email.get('body', '')
    segments = _re.split(r'\r?\n\r?\n', raw_body)
    classify_body = next(
        (s.strip() for s in segments if len(s.strip()) >= 5),
        raw_body[:200]
    )
    print(f"[negotiation:human] Classify body: {repr(classify_body[:150])}")

    # HUMAN NEGOTIATION: regex slot extraction — never hallucinate.
    # Matches 'HH:MM YYYY-MM-DD' (our documented format) in any order.
    # Also matches 'YYYY-MM-DD HH:MM' for flexibility.
    slot_pattern = _re.compile(
        r'(?:(?P<t1>\d{1,2}:\d{2})\s+(?P<d1>\d{4}-\d{2}-\d{2})'
        r'|(?P<d2>\d{4}-\d{2}-\d{2})\s+(?P<t2>\d{1,2}:\d{2}))'
    )
    regex_slots = []
    for m in slot_pattern.finditer(classify_body):
        date = m.group('d1') or m.group('d2')
        time = m.group('t1') or m.group('t2')
        # Zero-pad hour if needed (8:00 -> 08:00)
        if time and len(time) == 4:
            time = '0' + time
        if date and time:
            regex_slots.append({'date': date, 'time': time})
    print(f"[negotiation:human] Regex slots extracted: {regex_slots}")

    # LLM call for intent classification only — no slot extraction
    user_msg = (
        f"From: {email['from']} <{email['email']}>\n"
        f"Subject: {email['subject']}\n"
        f"Body:\n{classify_body}\n\n"
        f"Classify the intent only. Do NOT extract slots — they are handled separately. "
        f"Output JSON only."
    )
    response = ask_llm_chat(HUMAN_CLASSIFY_SYSTEM, user_msg)
    result   = safe_parse_json(response) or {}

    result.setdefault("type",             "fresh_request")
    result.setdefault("accepted_slot",    None)
    result.setdefault("proposed_date",    "")
    result.setdefault("duration_minutes", 60)
    result.setdefault("title",            "Meeting")
    result["is_agent"] = False

    # Always use regex slots — override whatever the LLM extracted
    result["proposed_slots"] = regex_slots

    # For accepted_slot, use regex result if LLM returned null
    ac = result.get("accepted_slot")
    if not isinstance(ac, dict) or not ac.get("date") or not ac.get("time"):
        # If classified as acceptance and we have one regex slot, use it
        if result["type"] == "acceptance" and len(regex_slots) == 1:
            result["accepted_slot"] = regex_slots[0]
        else:
            result["accepted_slot"] = None

    print(f"[negotiation:human] Classified as '{result['type']}' | "
          f"Slots extracted: {len(result['proposed_slots'])} | "
          f"Accepted slot: {result['accepted_slot']}")
    return result



def _human_compose_counter(
    original_email: dict,
    title: str,
    counter_slots: list[dict],
) -> str:
    """
    HUMAN NEGOTIATION: compose a natural-prose counter-proposal for a human.
    Slots are pre-formatted in Python and passed to the LLM as a locked string
    it must use verbatim - prevents Llama 3.2 hallucinating extra times.
    """
    # Build the slot phrase entirely in Python - LLM only writes the wrapper sentence
    dts = [
        datetime.strptime(f"{s['date']} {s['time']}", "%Y-%m-%d %H:%M")
        for s in counter_slots
    ]
    if len(dts) == 1:
        slot_phrase = _fmt_slot_human(dts[0])
    elif len(dts) == 2:
        slot_phrase = f"{_fmt_slot_human(dts[0])} or {_fmt_slot_human(dts[1])}"
    else:
        slot_phrase = ", ".join(_fmt_slot_human(d) for d in dts[:-1])
        slot_phrase += f", or {_fmt_slot_human(dts[-1])}"

    user_msg = (
        f"Original subject: {original_email['subject']}\n"
        f"Meeting title: {title}\n"
        f"You MUST use this exact phrase for the times (do not change it): '{slot_phrase}'\n"
        f"Write a warm 2-sentence reply saying the requested time is unavailable "
        f"and proposing the above times. Use the phrase exactly as given. Output JSON only."
    )
    response = ask_llm_chat(HUMAN_COUNTER_PROPOSAL_SYSTEM, user_msg)
    result   = safe_parse_json(response)
    if result and result.get("body"):
        # Verify the LLM didn't drop our slot phrase - if it did, use fallback
        body = result["body"]
        if any(_fmt_slot_human(d).split(" at ")[1] in body for d in dts):
            return body
    # Fallback: deterministic prose - guaranteed correct slots, no LLM involvement
    return (
        f"Unfortunately that time isn't available for {title}. "
        f"I'm free on {slot_phrase} - would any of those work for you?"
    )


def _human_compose_confirmation(
    original_email: dict,
    title: str,
    date: str,
    time_str: str,
) -> str:
    """
    HUMAN NEGOTIATION: compose a warm confirmation email for a human recipient.
    States the meeting title, date, and time clearly in plain prose.
    """
    readable_dt = _fmt_slot_human_long(
        datetime.strptime(f"{date} {time_str}", "%Y-%m-%d %H:%M")
    )
    user_msg = (
        f"Original subject: {original_email['subject']}\n"
        f"Meeting '{title}' confirmed for {readable_dt}.\n"
        f"Write a warm confirmation for a human recipient. Mention a calendar invite is attached. "
        f"Output JSON only."
    )
    response = ask_llm_chat(HUMAN_CONFIRMATION_SYSTEM, user_msg)
    result   = safe_parse_json(response)
    if result and result.get("body"):
        return result["body"]
    return (
        f"Confirmed - {title} is booked for {readable_dt}. "
        f"A calendar invite is attached for your records. Looking forward to it!"
    )


def _human_compose_rejection(original_email: dict, reason: str) -> str:
    """
    HUMAN NEGOTIATION: compose a polite rejection/failure email for a human.
    """
    user_msg = (
        f"Original subject: {original_email['subject']}\n"
        f"Reason: {reason}\n"
        f"Write a polite reply for a human recipient. Output JSON only."
    )
    response = ask_llm_chat(HUMAN_REJECTION_SYSTEM, user_msg)
    result   = safe_parse_json(response)
    if result and result.get("body"):
        return result["body"]
    return (
        "I'm sorry - we weren't able to find a mutually suitable time. "
        "Please feel free to reach out directly to reschedule."
    )


def _human_build_ics(title: str, date: str, time_str: str, duration: int, description: str = "") -> str:
    """
    HUMAN NEGOTIATION: build a plain-text ICS calendar invite string.
    Returned as a string so email_monitor / app.py can attach it.
    The human can import this directly into Outlook, Google Calendar, etc.
    """
    import uuid as _uuid
    start_dt  = datetime.strptime(f"{date} {time_str}", "%Y-%m-%d %H:%M")
    end_dt    = start_dt + timedelta(minutes=duration)
    dtstamp   = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    dtstart   = start_dt.strftime("%Y%m%dT%H%M%S")
    dtend     = end_dt.strftime("%Y%m%dT%H%M%S")
    uid       = str(_uuid.uuid4())

    def _esc(s: str) -> str:
        return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")

    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        "PRODID:-//Agent System//EN", "CALSCALE:GREGORIAN",
        "METHOD:REQUEST",
        "BEGIN:VEVENT",
        f"UID:{uid}", f"DTSTAMP:{dtstamp}",
        f"DTSTART:{dtstart}", f"DTEND:{dtend}",
        f"SUMMARY:{_esc(title)}",
    ]
    if description:
        lines.append(f"DESCRIPTION:{_esc(description)}")
    lines += ["END:VEVENT", "END:VCALENDAR", ""]
    return "\r\n".join(lines)


# -- HUMAN NEGOTIATION: reclassification --------------------------------------

def _human_reclassify_by_thread(
    email_type: str,
    thread_id: str,
    accepted_slot: dict | None,
) -> str:
    """
    HUMAN NEGOTIATION: simplified reclassification - no Rule 0 (no slot blocks).

    Rule H1 - rejection is always terminal, never overridden.

    Rule H2 - if LLM says acceptance and accepted_slot matches our last offer
               confirmed acceptance. If accepted_slot is missing or does not
               match, trust the LLM classification as-is.

    Rule H3 - fresh_request on an active thread -> counter_proposal.

    No mutual-tentative check (Rule 0) - humans do not book tentatives.
    """
    existing = _get_negotiation(thread_id)
    active   = existing and existing.get("status") == "active"

    if email_type == "rejection":
        return "rejection"

    if active:
        # Rule H2 - validate acceptance against offer history
        if email_type == "acceptance":
            if not accepted_slot:
                # LLM said acceptance but extracted no slot - trust it anyway
                print(f"[negotiation:human] Acceptance with no slot extracted - trusting classifier.")
                return "acceptance"

            # Check if accepted_slot matches anything we last offered
            last_offered = set()
            for entry in reversed(existing.get("history", [])):
                if entry.get("action") in ("tentative", "counter_proposal"):
                    for s in entry.get("slots_offered", [entry.get("slot", {})]):
                        if s and s.get("date") and s.get("time"):
                            last_offered.add((s["date"], s["time"]))
                    break

            if not last_offered:
                print(f"[negotiation:human] No offer history - trusting classifier: acceptance.")
                return "acceptance"

            key = (accepted_slot["date"], accepted_slot["time"])
            if key in last_offered:
                print(f"[negotiation:human] Accepted slot {key} matches our offer -> acceptance.")
                return "acceptance"
            else:
                print(f"[negotiation:human] Accepted slot {key} not in last offer "
                      f"{last_offered} -> counter_proposal.")
                return "counter_proposal"

        # Rule H3
        if email_type == "fresh_request":
            print(f"[negotiation:human] Reclassified fresh_request -> counter_proposal "
                  f"(active thread: {thread_id!r})")
            return "counter_proposal"

    return email_type


# -- HUMAN NEGOTIATION: core handlers -----------------------------------------

def _human_handle_fresh_request(
    email: dict,
    classification: dict,
    thread_id: str,
) -> dict:
    """
    HUMAN NEGOTIATION: handle the opening email from a human.
    If their proposed slot is free -> book tentative, reply accepting with
    a request for their confirmation.
    If busy -> counter-propose with natural prose listing alternatives.
    """
    title    = classification.get("title", "Meeting")
    duration = int(classification.get("duration_minutes") or 60)
    slots    = classification.get("proposed_slots", [])

    negotiation = {
        "thread_id":    thread_id,
        "title":        title,
        "duration":     duration,
        "initiator":    email["email"],
        "rounds":       0,
        "status":       "active",
        "history":      [],
        "created_at":   datetime.now().isoformat(),
        "human_thread": True,  # HUMAN NEGOTIATION: marks this as a human thread
    }

    for slot in slots:
        date, time_str = slot.get("date", ""), slot.get("time", "")
        if not date or not time_str:
            continue
        avail = check_availability(date, time_str, duration)
        if avail["available"]:
            start_dt = datetime.strptime(f"{date} {time_str}", "%Y-%m-%d %H:%M")
            end_dt   = start_dt + timedelta(minutes=duration)
            event    = {
                "title":            title,
                "duration_minutes": duration,
                "description":      f"Meeting with {email['from']}",
                "location":         "",
            }
            tent        = book_tentative(event, start_dt, end_dt)
            agreed_slot = {"date": date, "time": time_str}

            # Reply to human: propose the slot back in natural prose, ask for confirmation
            body = _human_compose_tentative_request(email, title, agreed_slot)

            negotiation["rounds"]   = 1
            negotiation["entry_id"] = tent.get("entry_id")
            negotiation["history"].append({
                "round":    1,
                "action":   "tentative",
                "slot":     agreed_slot,
                "entry_id": tent.get("entry_id"),
            })
            _upsert_negotiation(thread_id, negotiation)

            print(f"[negotiation:human] Tentative booked at {date} {time_str} - "
                  f"awaiting human confirmation.")
            return {
                "action":    "tentative",
                "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
                "slot":      agreed_slot,
                "entry_id":  tent.get("entry_id"),
                "thread_id": thread_id,
                "ics":       None,  # No ICS until human confirms
            }

    # Proposed slot(s) are busy - build a human-readable counter-proposal
    return _human_build_counter_proposal(email, classification, negotiation, thread_id, is_first=True)


def _human_handle_counter_proposal(
    email: dict,
    classification: dict,
    thread_id: str,
) -> dict:
    """
    HUMAN NEGOTIATION: handle a counter-proposal from a human.
    Checks their proposed slots against the calendar.
    If one is free -> book tentative, reply proposing it back in natural prose.
    If all busy -> counter with fresh slots.
    """
    negotiation = _get_negotiation(thread_id)

    if not negotiation:
        # Human initiated the chain - no record yet (same pattern as agent initiator)
        print(f"[negotiation:human] No record for {thread_id!r} - creating human initiator record.")
        negotiation = {
            "thread_id":    thread_id,
            "title":        classification.get("title", "Meeting"),
            "duration":     int(classification.get("duration_minutes") or 60),
            "initiator":    email["email"],
            "rounds":       0,
            "status":       "active",
            "history":      [],
            "entry_id":     None,
            "created_at":   datetime.now().isoformat(),
            "human_thread": True,  # HUMAN NEGOTIATION: marks this as a human thread
        }
        _upsert_negotiation(thread_id, negotiation)

    negotiation["rounds"] = negotiation.get("rounds", 0) + 1

    if negotiation["rounds"] > MAX_ROUNDS:
        return _human_handle_failure(email, negotiation, thread_id, reason="max_rounds")

    slots    = classification.get("proposed_slots", [])
    duration = negotiation.get("duration", 60)

    # HUMAN NEGOTIATION: sort slots chronologically so the human's newest
    # proposal is checked first. When a human says 'X doesn't work, can we do Y',
    # the classifier extracts both. Sorting ensures we try Y before X.
    def _slot_dt(s):
        try:
            return datetime.strptime(f"{s.get('date','')} {s.get('time','')}", "%Y-%m-%d %H:%M")
        except ValueError:
            return datetime.max
    slots = sorted(slots, key=_slot_dt)

    # Build a set of slots we previously offered so we can identify
    # which slots in this email are the human's NEW proposals vs ones they just rejected.
    # Build set of slots from MOST RECENT offer only.
    # Using full history caused a bug where if the agent offered slot X in Round 1
    # and the human asks for X in Round 2, it would be excluded as 'previously offered'.
    # Using only the last offer correctly identifies what the human is currently rejecting.
    last_offered = set()
    for entry in reversed(negotiation.get("history", [])):
        if entry.get("action") in ("tentative", "counter_proposal"):
            for s in entry.get("slots_offered", [entry.get("slot", {})]):
                if s and s.get("date") and s.get("time"):
                    last_offered.add((s["date"], s["time"]))
            break  # Only the most recent offer

    print(f"[negotiation:human] Last offered: {last_offered}")
    print(f"[negotiation:human] All extracted slots: {slots}")

    # HUMAN NEGOTIATION: two-pass slot checking.
    # Pass 1 - try only the human's NEW proposals (not in last offer).
    # Pass 1b - if human_proposed is empty and ALL extracted slots are from
    #           last_offered, the human is implicitly accepting one of our slots.
    #           Reroute to acceptance handler rather than counter-proposing.
    # Pass 2 - if no new proposals are free, search the calendar for alternatives.
    human_proposed = [
        s for s in slots
        if (s.get("date"), s.get("time")) not in last_offered
        and s.get("date") and s.get("time")
    ]

    print(f"[negotiation:human] Human proposed (new slots only): {human_proposed}")

    # Pass 1b - implicit acceptance check
    if not human_proposed and slots:
        all_in_last_offered = all(
            (s.get("date"), s.get("time")) in last_offered
            for s in slots if s.get("date") and s.get("time")
        )
        if all_in_last_offered:
            accepted = slots[0]
            print(f"[negotiation:human] All slots from last offer - "
                  f"treating as implicit acceptance of {accepted}.")
            return _human_handle_acceptance(email, thread_id, accepted)

    for slot in human_proposed:
        date, time_str = slot.get("date", ""), slot.get("time", "")
        avail = check_availability(date, time_str, duration)
        if avail["available"]:
            start_dt = datetime.strptime(f"{date} {time_str}", "%Y-%m-%d %H:%M")
            end_dt   = start_dt + timedelta(minutes=duration)
            event    = {
                "title":            negotiation["title"],
                "duration_minutes": duration,
                "description":      f"Meeting with {email['from']}",
                "location":         "",
            }
            tent        = book_tentative(event, start_dt, end_dt)
            agreed_slot = {"date": date, "time": time_str}

            # Reply to human: propose this slot in natural prose
            body = _human_compose_tentative_request(email, negotiation["title"], agreed_slot)

            negotiation["entry_id"] = tent.get("entry_id")
            negotiation["history"].append({
                "round":    negotiation["rounds"],
                "action":   "tentative",
                "slot":     agreed_slot,
                "entry_id": tent.get("entry_id"),
            })
            _upsert_negotiation(thread_id, negotiation)

            print(f"[negotiation:human] Tentative booked at {date} {time_str} - "
                  f"awaiting human confirmation.")
            return {
                "action":    "tentative",
                "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
                "slot":      agreed_slot,
                "entry_id":  tent.get("entry_id"),
                "thread_id": thread_id,
                "ics":       None,
            }

    return _human_build_counter_proposal(email, classification, negotiation, thread_id, is_first=False)


def _human_build_counter_proposal(
    email: dict,
    classification: dict,
    negotiation: dict,
    thread_id: str,
    is_first: bool,
) -> dict:
    """
    HUMAN NEGOTIATION: build a natural-prose counter-proposal for a human.
    Finds free slots using the same availability helpers as the agent path,
    but formats them as readable prose - no slot blocks, no bullet lists.
    """
    duration = negotiation.get("duration", int(classification.get("duration_minutes") or 60))

    if is_first:
        negotiation["rounds"] = negotiation.get("rounds", 0) + 1

    # Determine reference date from what the human proposed
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
        counter_slots = [{"date": proposed_date, "time": t} for t in free_today[:MAX_COUNTER_SLOTS]]
    else:
        future_slots = get_free_slots_next_n_days(proposed_date, duration)
        if not future_slots:
            return _human_handle_failure(email, negotiation, thread_id, reason="no_availability")
        counter_slots = []
        for date, times in future_slots.items():
            for t in times:
                counter_slots.append({"date": date, "time": t})
                if len(counter_slots) >= MAX_COUNTER_SLOTS:
                    break
            if len(counter_slots) >= MAX_COUNTER_SLOTS:
                break

    # Compose natural prose - slots written inline, no machine blocks appended
    body = _human_compose_counter(email, negotiation.get("title", "Meeting"), counter_slots)

    negotiation["history"] = negotiation.get("history", [])
    negotiation["history"].append({
        "round":         negotiation["rounds"],
        "action":        "counter_proposal",
        "slots_offered": counter_slots,
    })
    negotiation["status"] = "active"
    _upsert_negotiation(thread_id, negotiation)

    print(f"[negotiation:human] Counter-proposal sent with {len(counter_slots)} slot(s) "
          f"(prose only, no slot block).")
    return {
        "action":        "counter_proposal",
        "reply":         {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "slots_offered": counter_slots,
        "thread_id":     thread_id,
        "ics":           None,
    }


def _human_handle_acceptance(email: dict, thread_id: str, accepted_slot: dict | None) -> dict:
    """
    HUMAN NEGOTIATION: handle a human accepting a proposed slot.
    Promotes the tentative calendar entry to confirmed.
    Returns an ICS string so the caller can attach it to the reply email.
    No Rule 0 - convergence is via Rule H2 in _human_reclassify_by_thread().
    """
    negotiation = _get_negotiation(thread_id)
    duration    = (negotiation or {}).get("duration", 60)
    title       = (negotiation or {}).get("title", "Meeting")
    entry_id    = (negotiation or {}).get("entry_id")

    # Prefer our own tentative slot from history over what the LLM extracted
    tentative_slot = None
    for entry in reversed((negotiation or {}).get("history", [])):
        if entry.get("action") == "tentative" and entry.get("slot"):
            tentative_slot = entry["slot"]
            break

    slot_to_confirm = tentative_slot or accepted_slot

    if not slot_to_confirm:
        print(f"[negotiation:human] Acceptance received but no slot to confirm - "
              f"re-opening with fresh slots.")
        if negotiation:
            negotiation["status"] = "active"
            negotiation["rounds"] = negotiation.get("rounds", 0) + 1
            _upsert_negotiation(thread_id, negotiation)
        return _human_build_counter_proposal(
            email,
            {"proposed_slots": [], "duration_minutes": duration, "title": title},
            negotiation or {"title": title, "duration": duration, "history": [], "rounds": 1},
            thread_id,
            is_first=False,
        )

    # Verify slot is still free (race condition guard - same logic as agent path)
    avail     = check_availability(slot_to_confirm["date"], slot_to_confirm["time"], duration)
    conflicts = avail.get("conflicts", [])
    only_our_tentative = (
        not avail["available"] and
        all(c.get("tentative") and title.lower() in c.get("title", "").lower()
            for c in conflicts)
    )

    if avail["available"] or only_our_tentative:
        if entry_id:
            confirm_status = confirm_tentative(entry_id, title)
            print(f"[negotiation:human] {confirm_status}")
        else:
            from calendar_agent import calendar_agent as _cal
            _cal(f"{title} with {email['from']} on {slot_to_confirm['date']} "
                 f"at {slot_to_confirm['time']} for {duration} minutes")

        body = _human_compose_confirmation(
            email, title, slot_to_confirm["date"], slot_to_confirm["time"]
        )
        ics = _human_build_ics(
            title       = title,
            date        = slot_to_confirm["date"],
            time_str    = slot_to_confirm["time"],
            duration    = duration,
            description = f"Meeting with {email['from']}",
        )
        _close_negotiation(thread_id, "confirmed")
        print(f"[negotiation:human] Confirmed '{title}' at "
              f"{slot_to_confirm['date']} {slot_to_confirm['time']} - ICS generated.")

        return {
            "action":    "confirmed",
            "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
            "slot":      slot_to_confirm,
            "thread_id": thread_id,
            "ics":       ics,  # HUMAN NEGOTIATION: attach this to the outgoing email
        }

    else:
        # Slot now taken - cancel tentative and re-open
        if entry_id:
            cancel_tentative(entry_id, title)
        print(f"[negotiation:human] Tentative slot now conflicts - re-opening negotiation.")
        if negotiation:
            negotiation["status"]   = "active"
            negotiation["entry_id"] = None
            negotiation["rounds"]   = negotiation.get("rounds", 0) + 1
            _upsert_negotiation(thread_id, negotiation)

        ref_date = datetime.now().strftime("%Y-%m-%d")
        free     = get_all_free_slots(ref_date, duration)
        fresh    = ([{"date": ref_date, "time": t} for t in free[:MAX_COUNTER_SLOTS]]
                    if free else [])
        if not fresh:
            future = get_free_slots_next_n_days(ref_date, duration)
            for d, times in future.items():
                for t in times:
                    fresh.append({"date": d, "time": t})
                    if len(fresh) >= MAX_COUNTER_SLOTS:
                        break
                if len(fresh) >= MAX_COUNTER_SLOTS:
                    break

        body = _human_compose_counter(email, title, fresh)
        return {
            "action":        "counter_proposal",
            "reply":         {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
            "slots_offered": fresh,
            "thread_id":     thread_id,
            "ics":           None,
        }


def _human_handle_rejection(email: dict, thread_id: str) -> dict:
    """
    HUMAN NEGOTIATION: handle a human declining to meet.
    Cancels any tentative and closes the thread.
    """
    negotiation = _get_negotiation(thread_id)
    entry_id    = (negotiation or {}).get("entry_id")
    title       = (negotiation or {}).get("title", "Meeting")

    if entry_id:
        print(f"[negotiation:human] Human rejected - cancelling tentative.")
        cancel_tentative(entry_id, title)

    body = _human_compose_rejection(email, "The human has declined to proceed.")
    _close_negotiation(thread_id, "rejected")

    return {
        "action":    "rejected",
        "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "thread_id": thread_id,
        "ics":       None,
    }


def _human_handle_failure(
    email: dict,
    negotiation: dict,
    thread_id: str,
    reason: str,
) -> dict:
    """
    HUMAN NEGOTIATION: terminate a negotiation that has exceeded MAX_ROUNDS
    or has no available slots. Cancels tentative and sends a polite apology.
    """
    entry_id = negotiation.get("entry_id")
    title    = negotiation.get("title", "Meeting")

    if entry_id:
        print(f"[negotiation:human] Negotiation failed ({reason}) - cancelling tentative.")
        cancel_tentative(entry_id, title)

    reasons = {
        "max_rounds":      "We were unable to find a mutually agreeable time after several attempts.",
        "no_availability": "No available slots exist in the coming days.",
    }
    body = _human_compose_rejection(email, reasons.get(reason, "Negotiation could not be completed."))

    negotiation["history"] = negotiation.get("history", [])
    negotiation["history"].append({"action": "failed", "reason": reason})
    _upsert_negotiation(thread_id, negotiation)
    _close_negotiation(thread_id, "failed")

    return {
        "action":    "failed",
        "reply":     {"to": email["email"], "subject": f"Re: {email['subject']}", "body": body},
        "reason":    reason,
        "thread_id": thread_id,
        "ics":       None,
        "escalate":  True,
    }


# -- HUMAN NEGOTIATION: main entry point --------------------------------------

def negotiate_with_human(email: dict) -> dict | None:
    """
    HUMAN NEGOTIATION: drop-in parallel to negotiate() for human senders.

    Call this instead of negotiate() when email_monitor detects the sender
    is not an agent (is_agent_email() returns False).

    Uses independent classification, reclassification, handlers, and composers.
    The agent-to-agent negotiate() path is completely untouched.

    Returns the same dict shape as negotiate() so email_monitor can handle
    both paths identically - with one addition: "ics" key contains a calendar
    invite string on confirmed actions (None otherwise).

    Closed thread guard is identical to the agent path.
    """
    thread_id = _extract_thread_id(email.get("subject", ""))

    # Guard: ignore emails on already-closed threads (same as agent path)
    existing = _get_negotiation(thread_id)
    if existing and existing.get("status") in ("confirmed", "rejected", "failed"):
        print(f"[negotiation:human] Thread {thread_id!r} already closed "
              f"({existing['status']}) - ignoring email.")
        return None

    classification = _human_classify_email(email)
    email_type     = _human_reclassify_by_thread(
        classification["type"],
        thread_id,
        classification.get("accepted_slot"),
    )
    classification["type"] = email_type

    current_round = (existing or {}).get("rounds", 0) + 1
    print(
        f"[negotiation:human] Thread: {thread_id!r} | Type: {email_type} | "
        f"Round: {current_round}"
    )

    if email_type == "acceptance":
        return _human_handle_acceptance(email, thread_id, classification.get("accepted_slot"))
    if email_type == "rejection":
        return _human_handle_rejection(email, thread_id)
    if email_type == "fresh_request" and not existing:
        return _human_handle_fresh_request(email, classification, thread_id)
    if email_type in ("counter_proposal", "fresh_request"):
        return _human_handle_counter_proposal(email, classification, thread_id)

    print(f"[negotiation:human] Unhandled type '{email_type}' - requires manual review.")
    return None