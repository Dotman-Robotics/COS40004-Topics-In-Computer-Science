import json
import re
from ollama_client import ask_llm_chat



def _extract_first_json(text: str) -> str:
    text = re.sub(r"```(?:json)?", "", text).strip()

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


def _safe_parse(text: str) -> dict | None:
    cleaned = _extract_first_json(text)
    if not cleaned:
        return None
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        fixed = cleaned.replace(": true", ': true').replace(": false", ': false')
        try:
            return json.loads(fixed)
        except json.JSONDecodeError:
            return None


SINGLE_EMAIL_SYSTEM = """You are an email summarizer. Output ONLY a JSON object. No explanation, no markdown, no backticks.

Required fields:
- "summary": 1-2 sentence summary of what the email is about and what it needs
- "sender": sender's name as a string
- "category": exactly one of: "meeting_request", "action_required", "information", "reply_needed", "other"
- "action_items": array of strings describing specific tasks needed, or empty array []
- "meeting_details": object with keys "detected" (true/false), "date" (YYYY-MM-DD or ""), "time" (HH:MM or ""), "title" (string or "")
- "priority": exactly one of: "high", "medium", "low"

Rules:
- If the email asks for a meeting, set category to "meeting_request" and populate meeting_details
- If the email requires a reply or action, set appropriate category and list action_items
- If date/time are not explicitly stated, leave them as empty string ""
- priority is "high" if urgent or time-sensitive, "low" if informational only"""


BATCH_SUMMARY_SYSTEM = """You are an email inbox summarizer. Output ONLY a JSON object. No explanation, no markdown, no backticks.

You will receive a numbered list of pre-summarized emails. Organize them into a digest.

Required output format:
{
  "total_emails": <integer>,
  "digest": {
    "meeting_requests": [{"from": "...", "subject": "...", "summary": "...", "date": "...", "time": "..."}],
    "action_required": [{"from": "...", "subject": "...", "action": "..."}],
    "information": [{"from": "...", "subject": "...", "summary": "..."}],
    "other": [{"from": "...", "subject": "...", "summary": "..."}]
  },
  "overview": "2-3 sentence summary of the overall inbox state and what needs attention"
}

Place each email in exactly one category based on its type. Do not leave any email out."""


SESSION_SUMMARY_SYSTEM = """You are an email session summarizer. Output ONLY a JSON object. No explanation, no markdown, no backticks.

You will receive a log of emails processed by an automated monitor. Summarize what happened.

Required output format:
{
  "session_overview": "2-3 sentence summary of what happened during the session",
  "meetings_booked": [{"title": "...", "date": "...", "time": "...", "with": "..."}],
  "meetings_declined": [{"title": "...", "date": "...", "reason": "conflict", "alternatives_offered": ["..."]}],
  "clarifications_requested": [{"from": "...", "subject": "..."}],
  "non_meeting_emails": [{"from": "...", "subject": "...", "summary": "..."}],
  "attention_needed": ["list anything requiring manual follow-up, or empty array"]
}

Be specific. Use names, titles, and dates from the log rather than generic placeholders."""

def summarize_email(email: dict) -> dict:
    body = (email.get("body") or "").strip()
    if not body:
        body = "(no body)"

    user_msg = (
        f"From: {email.get('from', 'Unknown')} <{email.get('email', '')}>\n"
        f"Subject: {email.get('subject', '(no subject)')}\n"
        f"Received: {email.get('received_time', '')}\n"
        f"Body:\n{body}\n\n"
        f"Summarize this email. Output JSON only."
    )

    response = ask_llm_chat(SINGLE_EMAIL_SYSTEM, user_msg)
    result   = _safe_parse(response)

    if result is None:
        print(f"[summarizer] Failed to parse single-email response:\n{response[:300]}")
        return _fallback_summary(email)

    result.setdefault("summary",      email.get("subject", "No summary available"))
    result.setdefault("sender",       email.get("from", "Unknown"))
    result.setdefault("category",     "other")
    result.setdefault("action_items", [])
    result.setdefault("priority",     "medium")

    md = result.get("meeting_details") # Enforce meeting_details structure
    if not isinstance(md, dict):
        md = {}
    result["meeting_details"] = {
        "detected": bool(md.get("detected", False)),
        "date":     str(md.get("date",  "") or ""),
        "time":     str(md.get("time",  "") or ""),
        "title":    str(md.get("title", "") or ""),
    }

    valid_categories = {"meeting_request", "action_required", "information", "reply_needed", "other"}
    if result["category"] not in valid_categories:
        result["category"] = "other"

    if result["priority"] not in {"high", "medium", "low"}: #priority check
        result["priority"] = "medium"

    if not isinstance(result["action_items"], list):
        result["action_items"] = []
    result["action_items"] = [str(a) for a in result["action_items"] if a]

    return result


def _fallback_summary(email: dict) -> dict:
    return {
        "summary":      f"Email from {email.get('from', 'Unknown')}: {email.get('subject', '(no subject)')}",
        "sender":       email.get("from", "Unknown"),
        "category":     "other",
        "action_items": [],
        "meeting_details": {"detected": False, "date": "", "time": "", "title": ""},
        "priority":     "medium",
    }

def summarize_inbox_batch(emails: list[dict]) -> dict:
    if not emails:
        return {
            "total_emails": 0,
            "digest": {"meeting_requests": [], "action_required": [], "information": [], "other": []},
            "overview": "No emails to summarize.",
        }

    email_entries = ""
    for i, email in enumerate(emails, 1):
        summary = email.get("summary") or summarize_email(email) 
        if isinstance(summary, dict):
            summary_text = summary.get("summary", "")
            category     = summary.get("category", "other")
            action_items = summary.get("action_items", [])
            meeting      = summary.get("meeting_details", {})
        else:
            summary_text = str(summary)
            category     = "other"
            action_items = []
            meeting      = {}

        entry = (
            f"\n--- Email {i} ---\n"
            f"From: {email.get('from', 'Unknown')} <{email.get('email', '')}>\n"
            f"Subject: {email.get('subject', '(no subject)')}\n"
            f"Category: {category}\n"
            f"Summary: {summary_text}\n"
        )
        if action_items:
            entry += f"Action items: {', '.join(action_items)}\n"
        if meeting.get("detected"):
            entry += f"Meeting: {meeting.get('title','')} on {meeting.get('date','')} at {meeting.get('time','')}\n"
        email_entries += entry

    user_msg = (
        f"Organize these {len(emails)} emails into a digest. Output JSON only.\n"
        f"{email_entries}"
    )

    response = ask_llm_chat(BATCH_SUMMARY_SYSTEM, user_msg)
    result   = _safe_parse(response)

    if result is None:
        print(f"[summarizer] Failed to parse batch response:\n{response[:300]}")
        return _batch_fallback(emails)

    result["total_emails"] = len(emails)

    digest = result.setdefault("digest", {})
    for key in ("meeting_requests", "action_required", "information", "other"):
        digest.setdefault(key, [])

    result.setdefault("overview", f"{len(emails)} email(s) processed.")
    return result


def _batch_fallback(emails: list[dict]) -> dict:
    digest = {"meeting_requests": [], "action_required": [], "information": [], "other": []}

    for email in emails:
        summary  = email.get("summary") or summarize_email(email)
        category = summary.get("category", "other") if isinstance(summary, dict) else "other"
        entry = {
            "from":    email.get("from", "Unknown"),
            "subject": email.get("subject", "(no subject)"),
            "summary": summary.get("summary", "") if isinstance(summary, dict) else str(summary),
        }
        if category == "meeting_request":
            md = summary.get("meeting_details", {}) if isinstance(summary, dict) else {}
            entry["date"] = md.get("date", "")
            entry["time"] = md.get("time", "")
            digest["meeting_requests"].append(entry)
        elif category == "action_required":
            items = summary.get("action_items", []) if isinstance(summary, dict) else []
            entry["action"] = items[0] if items else ""
            digest["action_required"].append(entry)
        elif category == "information":
            digest["information"].append(entry)
        else:
            digest["other"].append(entry)

    return {
        "total_emails": len(emails),
        "digest":       digest,
        "overview":     f"{len(emails)} email(s) processed.",
    }

def summarize_monitor_session(session_log: list[dict]) -> dict:
    if not session_log:
        return {
            "session_overview":         "No emails were processed during this session.",
            "meetings_booked":          [],
            "meetings_declined":        [],
            "clarifications_requested": [],
            "non_meeting_emails":       [],
            "attention_needed":         [],
        }

    log_text = ""
    for i, entry in enumerate(session_log, 1):
        status  = entry.get("status", "unknown")
        email   = entry.get("email",   {})
        meeting = entry.get("meeting") or {}
        summary = entry.get("summary") or {}

        log_text += (
            f"\n--- Entry {i} ---\n"
            f"Status: {status}\n"
            f"From: {email.get('from', '')} <{email.get('email', '')}>\n"
            f"Subject: {email.get('subject', '')}\n"
        )

        if isinstance(summary, dict) and summary.get("summary"):
            log_text += f"Summary: {summary['summary']}\n"
            if summary.get("action_items"):
                log_text += f"Action items: {', '.join(summary['action_items'])}\n"

        if meeting.get("is_meeting_request"):
            log_text += (
                f"Meeting title: {meeting.get('title', '')}\n"
                f"Date: {meeting.get('date', '')}  Time: {meeting.get('time', '')}\n"
                f"Requested by: {meeting.get('requester_name', email.get('from', ''))}\n"
            )

        if status == "confirmed":
            cal = entry.get("calendar", {})
            log_text += f"Outcome: Meeting was booked. Calendar status: {cal.get('outlook_status', '')}\n"

        elif status == "declined":
            conflicts    = [c.get("title", "") for c in entry.get("conflicts", [])]
            alternatives = entry.get("alternatives", [])
            log_text += (
                f"Outcome: Declined due to conflict with: {', '.join(conflicts) or 'unknown'}\n"
                f"Alternatives offered: {', '.join(alternatives) or 'none'}\n"
            )

        elif status == "clarification_needed":
            log_text += "Outcome: Clarification email sent — missing date or time.\n"

        elif status == "non_meeting":
            log_text += "Outcome: Not a meeting request, logged only.\n"

    user_msg = (
        f"Summarize this email monitor session log. Output JSON only.\n"
        f"{log_text}"
    )

    response = ask_llm_chat(SESSION_SUMMARY_SYSTEM, user_msg)
    result   = _safe_parse(response)

    if result is None:
        print(f"[summarizer] Failed to parse session response:\n{response[:300]}")
        return _session_fallback(session_log)

    result.setdefault("session_overview",         "Session complete.")
    result.setdefault("meetings_booked",          [])
    result.setdefault("meetings_declined",        [])
    result.setdefault("clarifications_requested", [])
    result.setdefault("non_meeting_emails",       [])
    result.setdefault("attention_needed",         [])

    return result


def _session_fallback(session_log: list[dict]) -> dict:
    booked   = [e for e in session_log if e.get("status") == "confirmed"]
    declined = [e for e in session_log if e.get("status") == "declined"]
    clarify  = [e for e in session_log if e.get("status") == "clarification_needed"]
    other    = [e for e in session_log if e.get("status") == "non_meeting"]

    return {
        "session_overview": (
            f"Processed {len(session_log)} email(s): "
            f"{len(booked)} booked, {len(declined)} declined, "
            f"{len(clarify)} needed clarification, {len(other)} non-meeting."
        ),
        "meetings_booked": [
            {
                "title": e.get("meeting", {}).get("title", "Meeting"),
                "date":  e.get("meeting", {}).get("date",  ""),
                "time":  e.get("meeting", {}).get("time",  ""),
                "with":  e.get("email",   {}).get("from",  ""),
            }
            for e in booked
        ],
        "meetings_declined": [
            {
                "title":               e.get("meeting", {}).get("title", "Meeting"),
                "date":                e.get("meeting", {}).get("date",  ""),
                "reason":              "conflict",
                "alternatives_offered": e.get("alternatives", []),
            }
            for e in declined
        ],
        "clarifications_requested": [
            {
                "from":    e.get("email", {}).get("from",    ""),
                "subject": e.get("email", {}).get("subject", ""),
            }
            for e in clarify
        ],
        "non_meeting_emails": [
            {
                "from":    e.get("email", {}).get("from",    ""),
                "subject": e.get("email", {}).get("subject", ""),
                "summary": (e.get("summary") or {}).get("summary", "No summary available"),
            }
            for e in other
        ],
        "attention_needed": [],
    }

def print_email_summary(summary: dict, index: int | None = None) -> None:
    prefix = f"Email {index} — " if index else ""
    print(f"\n  {prefix}{summary.get('sender', 'Unknown')}")
    print(f"  Category : {summary.get('category', 'other').replace('_', ' ').title()}")
    print(f"  Priority : {summary.get('priority', 'medium').title()}")
    print(f"  Summary  : {summary.get('summary', '')}")

    actions = summary.get("action_items", [])
    if actions:
        print(f"  Actions  :")
        for action in actions:
            print(f"    • {action}")

    meeting = summary.get("meeting_details", {})
    if meeting.get("detected"):
        print(f"  Meeting  : {meeting.get('title', '')} on {meeting.get('date', '')} at {meeting.get('time', '')}")


def print_batch_summary(digest: dict) -> None:
    print(f"\n{'═'*55}")
    print(f"  INBOX SUMMARY — {digest.get('total_emails', 0)} email(s)")
    print(f"{'═'*55}")
    print(f"\n  {digest.get('overview', '')}")

    d = digest.get("digest", {})

    if d.get("meeting_requests"):
        print(f"\n  MEETING REQUESTS ({len(d['meeting_requests'])})")
        for m in d["meeting_requests"]:
            date_time = f"{m.get('date', '')} {m.get('time', '')}".strip()
            print(f"    • {m.get('from', '')} — {m.get('subject', '')}")
            if date_time:
                print(f"      When: {date_time}")
            print(f"      {m.get('summary', '')}")

    if d.get("action_required"):
        print(f"\n  ACTION REQUIRED ({len(d['action_required'])})")
        for a in d["action_required"]:
            print(f"    • {a.get('from', '')} — {a.get('subject', '')}")
            print(f"      Action: {a.get('action', '')}")

    if d.get("information"):
        print(f"\n  INFO ({len(d['information'])})")
        for info in d["information"]:
            print(f"    • {info.get('from', '')} — {info.get('subject', '')}")
            print(f"      {info.get('summary', '')}")

    if d.get("other"):
        print(f"\n  OTHER ({len(d['other'])})")
        for o in d["other"]:
            print(f"    • {o.get('from', '')} — {o.get('subject', '')}")

    print(f"\n{'═'*55}\n")


def print_session_summary(report: dict) -> None:
    print(f"\n{'═'*55}")
    print(f"  MONITOR SESSION SUMMARY")
    print(f"{'═'*55}")
    print(f"\n  {report.get('session_overview', '')}")

    booked = report.get("meetings_booked", [])
    if booked:
        print(f"\n  MEETINGS BOOKED ({len(booked)})")
        for m in booked:
            print(f"    • {m.get('title', 'Meeting')} with {m.get('with', '')}")
            print(f"      {m.get('date', '')} at {m.get('time', '')}")

    declined = report.get("meetings_declined", [])
    if declined:
        print(f"\n  MEETINGS DECLINED ({len(declined)})")
        for m in declined:
            alts = ", ".join(m.get("alternatives_offered", [])) or "none"
            print(f"    • {m.get('title', 'Meeting')} on {m.get('date', '')}")
            print(f"      Alternatives offered: {alts}")

    clarify = report.get("clarifications_requested", [])
    if clarify:
        print(f"\n  CLARIFICATIONS REQUESTED ({len(clarify)})")
        for c in clarify:
            print(f"    • {c.get('from', '')} — {c.get('subject', '')}")

    non_meetings = report.get("non_meeting_emails", [])
    if non_meetings:
        print(f"\n  OTHER EMAILS ({len(non_meetings)})")
        for e in non_meetings:
            print(f"    • {e.get('from', '')} — {e.get('subject', '')}")
            if e.get("summary"):
                print(f"      {e['summary']}")

    attention = report.get("attention_needed", [])
    if attention:
        print(f"\n  NEEDS ATTENTION")
        for item in attention:
            print(f"    • {item}")

    print(f"\n{'═'*55}\n")