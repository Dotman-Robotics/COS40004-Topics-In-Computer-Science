from ollama_client import ask_llm_chat

# ── System prompts ────────────────────────────────────────────────────────────

SINGLE_EMAIL_SYSTEM = """You are an email assistant. Summarize the email clearly and concisely.

Output JSON only. No explanation.

{
  "summary": "1-2 sentence summary of what the email is about",
  "sender": "sender name",
  "category": "meeting_request | action_required | information | reply_needed | other",
  "action_items": ["list of specific things that need to be done, or empty list"],
  "meeting_details": {
    "detected": true/false,
    "date": "YYYY-MM-DD or ''",
    "time": "HH:MM or ''",
    "title": "meeting title or ''"
  },
  "priority": "high | medium | low"
}"""


BATCH_SUMMARY_SYSTEM = """You are an email assistant. Summarize a batch of emails into an organized digest.

Output JSON only. No explanation.

{
  "total_emails": 0,
  "digest": {
    "meeting_requests": [
      {"from": "...", "subject": "...", "summary": "...", "date": "...", "time": "..."}
    ],
    "action_required": [
      {"from": "...", "subject": "...", "action": "..."}
    ],
    "information": [
      {"from": "...", "subject": "...", "summary": "..."}
    ],
    "other": [
      {"from": "...", "subject": "...", "summary": "..."}
    ]
  },
  "overview": "2-3 sentence overall summary of the inbox state"
}"""


SESSION_SUMMARY_SYSTEM = """You are an email assistant. Summarize a monitor session report.

Output JSON only. No explanation.

{
  "session_overview": "2-3 sentence summary of what happened during the session",
  "meetings_booked": [
    {"title": "...", "date": "...", "time": "...", "with": "..."}
  ],
  "meetings_declined": [
    {"title": "...", "date": "...", "reason": "conflict", "alternatives_offered": ["..."]}
  ],
  "clarifications_requested": [
    {"from": "...", "subject": "..."}
  ],
  "non_meeting_emails": [
    {"from": "...", "subject": "...", "summary": "..."}
  ],
  "attention_needed": ["list of anything that requires manual follow-up"]
}"""

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

def summarize_email(email: dict) -> dict: 
    user_msg = f"""From: {email.get('from', 'Unknown')} <{email.get('email', '')}>
Subject: {email.get('subject', '(no subject)')}
Received: {email.get('received_time', '')}
Body:
{email.get('body', '')}

JSON:"""

    response = ask_llm_chat(SINGLE_EMAIL_SYSTEM, user_msg)
    cleaned  = _extract_first_json(response)

    if not cleaned:
        return _fallback_summary(email)

    try:
        import json
        result = json.loads(cleaned)
        result.setdefault("summary",         email.get("subject", "No summary available"))
        result.setdefault("sender",          email.get("from", "Unknown"))
        result.setdefault("category",        "other")
        result.setdefault("action_items",    [])
        result.setdefault("meeting_details", {"detected": False, "date": "", "time": "", "title": ""})
        result.setdefault("priority",        "medium")
        return result
    except Exception:
        return _fallback_summary(email)


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
            "digest": {
                "meeting_requests": [],
                "action_required":  [],
                "information":      [],
                "other":            [],
            },
            "overview": "No emails to summarize.",
        }

    email_list = ""
    for i, email in enumerate(emails, 1):
        email_list += (
            f"\n--- Email {i} ---\n"
            f"From: {email.get('from', 'Unknown')} <{email.get('email', '')}>\n"
            f"Subject: {email.get('subject', '(no subject)')}\n"
            f"Received: {email.get('received_time', '')}\n"
            f"Body: {email.get('body', '')[:300]}\n"
        )

    user_msg = f"Summarize these {len(emails)} emails:\n{email_list}\n\nJSON:"
    response = ask_llm_chat(BATCH_SUMMARY_SYSTEM, user_msg)
    cleaned  = _extract_first_json(response)

    if not cleaned:
        return _batch_fallback(emails)

    try:
        import json
        result = json.loads(cleaned)
        result["total_emails"] = len(emails)
        return result
    except Exception:
        return _batch_fallback(emails)


def _batch_fallback(emails: list[dict]) -> dict:
    digest = {
        "meeting_requests": [],
        "action_required":  [],
        "information":      [],
        "other":            [],
    }
    for email in emails:
        summary = summarize_email(email)
        entry   = {
            "from":    email.get("from", "Unknown"),
            "subject": email.get("subject", "(no subject)"),
            "summary": summary["summary"],
        }
        category = summary.get("category", "other")
        if category == "meeting_request":
            entry["date"] = summary["meeting_details"].get("date", "")
            entry["time"] = summary["meeting_details"].get("time", "")
            digest["meeting_requests"].append(entry)
        elif category == "action_required":
            entry["action"] = summary["action_items"][0] if summary["action_items"] else ""
            digest["action_required"].append(entry)
        elif category == "information":
            digest["information"].append(entry)
        else:
            digest["other"].append(entry)

    return {
        "total_emails": len(emails),
        "digest":       digest,
        "overview":     f"{len(emails)} emails processed.",
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
        meeting = entry.get("meeting", {})

        log_text += (
            f"\n--- Entry {i} ---\n"
            f"Status: {status}\n"
            f"From: {email.get('from', '')} <{email.get('email', '')}>\n"
            f"Subject: {email.get('subject', '')}\n"
        )
        if meeting and meeting.get("is_meeting_request"):
            log_text += (
                f"Meeting: {meeting.get('title', '')} on "
                f"{meeting.get('date', '')} at {meeting.get('time', '')}\n"
            )
        if entry.get("conflicts"):
            log_text += f"Conflicts: {[c.get('title') for c in entry['conflicts']]}\n"
        if entry.get("alternatives"):
            log_text += f"Alternatives offered: {entry['alternatives']}\n"

    user_msg = f"Summarize this monitor session:\n{log_text}\n\nJSON:"

    response = ask_llm_chat(SESSION_SUMMARY_SYSTEM, user_msg)
    cleaned  = _extract_first_json(response)

    if not cleaned:
        return _session_fallback(session_log)

    try:
        import json
        return json.loads(cleaned)
    except Exception:
        return _session_fallback(session_log)


def _session_fallback(session_log: list[dict]) -> dict:
    booked       = [e for e in session_log if e.get("status") == "confirmed"]
    declined     = [e for e in session_log if e.get("status") == "declined"]
    clarify      = [e for e in session_log if e.get("status") == "clarification_needed"]
    non_meetings = [e for e in session_log if e.get("status") not in
                    ("confirmed", "declined", "clarification_needed")]

    return {
        "session_overview": (
            f"Processed {len(session_log)} email(s): "
            f"{len(booked)} booked, {len(declined)} declined, "
            f"{len(clarify)} needed clarification."
        ),
        "meetings_booked": [
            {
                "title": e.get("meeting", {}).get("title", "Meeting"),
                "date":  e.get("meeting", {}).get("date", ""),
                "time":  e.get("meeting", {}).get("time", ""),
                "with":  e.get("email",   {}).get("from", ""),
            }
            for e in booked
        ],
        "meetings_declined": [
            {
                "title":               e.get("meeting", {}).get("title", "Meeting"),
                "date":                e.get("meeting", {}).get("date", ""),
                "reason":              "conflict",
                "alternatives_offered": e.get("alternatives", []),
            }
            for e in declined
        ],
        "clarifications_requested": [
            {
                "from":    e.get("email", {}).get("from", ""),
                "subject": e.get("email", {}).get("subject", ""),
            }
            for e in clarify
        ],
        "non_meeting_emails": [
            {
                "from":    e.get("email", {}).get("from", ""),
                "subject": e.get("email", {}).get("subject", ""),
                "summary": "No summary available.",
            }
            for e in non_meetings
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
            print(f"    {action}")

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
            print(f"    • {m['from']} — {m['subject']}")
            if date_time:
                print(f"      When: {date_time}")
            print(f"      {m.get('summary', '')}")

    if d.get("action_required"):
        print(f"\n  ACTION REQUIRED ({len(d['action_required'])})")
        for a in d["action_required"]:
            print(f"    • {a['from']} — {a['subject']}")
            print(f"      Action: {a.get('action', '')}")

    if d.get("information"):
        print(f"\n  INFO ({len(d['information'])})")
        for info in d["information"]:
            print(f"    • {info['from']} — {info['subject']}")
            print(f"      {info.get('summary', '')}")

    if d.get("other"):
        print(f"\n  OTHER ({len(d['other'])})")
        for o in d["other"]:
            print(f"    • {o['from']} — {o['subject']}")

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
