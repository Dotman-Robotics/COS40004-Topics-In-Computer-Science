import time
import threading
import json
from datetime import datetime
from ollama_client import ask_llm_chat
from utils import extract_first_json, get_outlook_inbox
from negotiation_agent import negotiate, negotiate_with_human, is_agent_email, purge_active_negotiations
from rag_store import query_sender_history, index_session
from email_summarizer import summarize_email

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("[monitor] win32com not available.")

import os

POLL_INTERVAL_SECONDS = 30

def _get_account() -> str:
    """Read account at call time so --account CLI arg is always respected."""
    return os.environ.get("AGENT_ACCOUNT", "zoomertron@outlook.com")

_monitor_thread: threading.Thread | None = None
_stop_event     = threading.Event()
_session_log: list[dict] = []
_session_log_lock        = threading.Lock()

# ── Blacklist ─────────────────────────────────────────────────────────────────

_blacklist: set[str] = set()
_blacklist_lock      = threading.Lock()


def _load_blacklist() -> None:
    global _blacklist
    try:
        with open(f"blacklist_{_get_account().replace("@","_").replace(".","_")}.json", "r", encoding="utf-8") as f:
            data = json.load(f)
        with _blacklist_lock:
            _blacklist = {addr.lower().strip() for addr in data if addr.strip()}
        print(f"[monitor] Blacklist loaded: {len(_blacklist)} address(es).")
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"[monitor] Could not load blacklist: {e}")


def _save_blacklist() -> None:
    with _blacklist_lock:
        data = sorted(_blacklist)
    try:
        with open(f"blacklist_{_get_account().replace("@","_").replace(".","_")}.json", "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"[monitor] Could not save blacklist: {e}")


def add_to_blacklist(email_address: str) -> bool:
    addr = email_address.lower().strip()
    if not addr:
        return False
    with _blacklist_lock:
        if addr in _blacklist:
            return False
        _blacklist.add(addr)
    _save_blacklist()
    print(f"[monitor] Blacklisted: {addr}")
    return True


def remove_from_blacklist(email_address: str) -> bool:
    addr = email_address.lower().strip()
    with _blacklist_lock:
        if addr not in _blacklist:
            return False
        _blacklist.discard(addr)
    _save_blacklist()
    print(f"[monitor] Removed from blacklist: {addr}")
    return True


def get_blacklist() -> list[str]:
    with _blacklist_lock:
        return sorted(_blacklist)


def is_blacklisted(email_address: str) -> bool:
    with _blacklist_lock:
        return email_address.lower().strip() in _blacklist


_load_blacklist()



def _strip_quoted_reply(body: str) -> str:
    """
    Strip quoted reply content from an Outlook email body.
    Outlook reply chains include separator patterns before the quoted original.
    Keeping only the newest human-written paragraph prevents the date/slot
    classifier from picking up dates from the agent's previous emails.
    Returns up to 800 chars of the stripped text.

    HUMAN NEGOTIATION: critical for correct date extraction in multi-round threads.
    Also used for agent emails (harmless — agent emails don't contain these separators).
    """
    import re
    # Common Outlook quoted-text separators
    separators = [
        r"_{5,}",                        # ___________
        r"-{5,}\s*Original Message",      # ----- Original Message -----
        r"From:\s+\S+@\S+",              # From: someone@example.com
        r"On .{5,} wrote:",              # On Mon, 6 Jul 2026 ... wrote:
        r"-----\s*Forwarded",            # ----- Forwarded message -----
    ]
    pattern = "|".join(separators)
    match = re.search(pattern, body, re.IGNORECASE)
    if match:
        body = body[:match.start()].strip()
    return body[:800]


# ── Meeting detection ─────────────────────────────────────────────────────────
# Uses the category already assigned by email_summarizer where possible.
# Falls back to a direct LLM check only when the summary is unavailable.
#
# All detailed classification (fresh_request / counter_proposal / acceptance /
# rejection) is handled exclusively by negotiation_agent.classify_negotiation_email.
# This function is intentionally kept as a binary gate only.

MEETING_CATEGORIES = {"meeting_request", "reply_needed"}

# Domains that send automated/marketing emails - never route to negotiation agent.
_MARKETING_DOMAINS = {
    "zoom.us", "zoominfo.com", "mail.zoom.us", "no-reply.zoom.us",
    "notifications.zoom.us", "noreply.zoom.us",
    "mail.google.com", "accounts.google.com",
    "microsoft.com", "microsoftonline.com",
    "linkedin.com", "facebook.com", "twitter.com", "instagram.com",
    "mailchimp.com", "sendgrid.net", "constantcontact.com",
    "hubspot.com", "salesforce.com", "marketo.com",
}

_MARKETING_SUBJECT_PATTERNS = [
    "unsubscribe", "no-reply", "noreply", "newsletter", "notification",
    "booking page is ready", "free trial", "verify your", "confirm your email",
    "password reset", "security alert", "invoice", "receipt",
    "your account", "activate your", "welcome to",
]


def _is_marketing_email(email: dict) -> bool:
    """
    Detect automated/marketing emails that should never reach the negotiation agent.
    Checks sender domain and subject line patterns.
    """
    sender  = email.get("email", "").lower()
    subject = email.get("subject", "").lower()

    domain = sender.split("@")[-1] if "@" in sender else ""
    if any(domain == d or domain.endswith("." + d) for d in _MARKETING_DOMAINS):
        return True

    local = sender.split("@")[0] if "@" in sender else ""
    if any(p in local for p in ("noreply", "no-reply", "donotreply",
                                 "notifications", "mailer", "bounce",
                                 "postmaster", "automailer")):
        return True

    if any(p in subject for p in _MARKETING_SUBJECT_PATTERNS):
        return True

    return False


def _is_meeting_related(summary: dict, email: dict) -> bool:
    """
    Binary gate: should this email be passed to the negotiation agent?
    Rejects marketing/automated senders first, then checks summarizer category.
    Falls back to LLM only when category is 'other' and body contains datetime patterns.
    """
    if _is_marketing_email(email):
        print("[monitor] Marketing/automated email detected - skipping.")
        return False

    category = summary.get("category", "other")

    if category in MEETING_CATEGORIES:
        return True

    if category == "information":
        return False

    import re as _re
    body = email.get("body", "")
    has_datetime = bool(_re.search(
        r'\b(\d{1,2}:\d{2}|\d{4}-\d{2}-\d{2}|monday|tuesday|wednesday|thursday|friday|'
        r'tomorrow|next week|am\b|pm\b)\b',
        body, _re.IGNORECASE
    ))
    if not has_datetime:
        return False

    DETECTOR_SYSTEM = """You are an email classifier. Output JSON only. Never explain.
Determine if this is a PERSONAL email from a real human requesting, proposing, accepting,
or rejecting a specific meeting. Do NOT return true for marketing emails, newsletters,
automated notifications, or emails that merely mention scheduling features.
Output: {\"is_meeting_related\": true/false}"""

    user_msg = (
        f"From: {email['from']} <{email['email']}>\n"
        f"Subject: {email['subject']}\n"
        f"Body: {email['body']}\n\nJSON:"
    )
    response = ask_llm_chat(DETECTOR_SYSTEM, user_msg)
    cleaned  = extract_first_json(response)
    if not cleaned:
        return False
    try:
        return bool(json.loads(cleaned).get("is_meeting_related", False))
    except json.JSONDecodeError:
        return False


def _send_reply(reply: dict, ics: str | None = None) -> bool:
    """
    Send a reply dict {to, subject, body} via Outlook COM.

    ics parameter: HUMAN NEGOTIATION only.
    When provided (a plain-text ICS string), attaches it as a .ics file
    so the human recipient can import the confirmed event into their calendar.
    """
    try:
        ol_app = win32com.client.Dispatch("Outlook.Application")
        ol_ns  = ol_app.GetNameSpace("MAPI")
        mail   = ol_app.CreateItem(0)

        mail.Subject = reply["subject"]
        mail.Body    = reply["body"]
        mail.To      = reply["to"]

        try:
            account = ol_ns.Accounts.Item(_get_account())
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
        except Exception as e:
            print(f"[monitor] Could not set sender account: {e}")

        # HUMAN NEGOTIATION: attach ICS file on confirmed bookings
        if ics:
            import tempfile
            import os as _os
            with tempfile.NamedTemporaryFile(
                mode="w", suffix=".ics", delete=False, encoding="utf-8"
            ) as tmp:
                tmp.write(ics)
                tmp_path = tmp.name
            try:
                mail.Attachments.Add(tmp_path)
                print(f"[monitor] ICS calendar invite attached.")
            except Exception as e:
                print(f"[monitor] Could not attach ICS: {e}")
            finally:
                try:
                    _os.unlink(tmp_path)
                except Exception:
                    pass

        mail.Display()
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


# ── Core email processor ──────────────────────────────────────────────────────

def process_email(msg) -> dict | None:
    """
    Process a single Outlook message object through the full pipeline:

      1. Read fields from COM object
      2. Blacklist check  →  drop immediately if blocked
      3. Summarize        →  email_summarizer (category, priority, action items)
      4. Meeting gate     →  binary check using summary category + fallback LLM
      5. Negotiate        →  negotiation_agent owns ALL scheduling logic
      6. Send reply       →  _send_reply via Outlook COM
      7. Log              →  append to session log under lock

    The monitor intentionally contains NO scheduling logic of its own.
    It is a routing and transport layer only.
    """
    # ── 1. Read message ───────────────────────────────────────────────────────
    try:
        email = {
            "from":          msg.SenderName,
            "email":         msg.SenderEmailAddress,
            "subject":       msg.Subject,
            "body":          _strip_quoted_reply(msg.Body[:3000]),  # Strip quoted thread before classifying — prevents date extraction from old messages
            "received_time": str(msg.ReceivedTime),
        }
    except Exception as e:
        print(f"[monitor] Could not read message: {e}")
        return None

    print(f"\n[monitor] ── New email from {email['from']}: '{email['subject']}'")

    # ── 2. Blacklist check ────────────────────────────────────────────────────
    if is_blacklisted(email["email"]):
        print(f"[monitor] Blocked — {email['email']} is blacklisted.")
        _mark_as_read(msg)
        result = {
            "status":  "blacklisted",
            "email":   email,
            "summary": None,
        }
        with _session_log_lock:
            _session_log.append(result)
        return result

    # ── 3. Summarize ──────────────────────────────────────────────────────────
    summary = summarize_email(email)
    print(f"[monitor] Summary:  {summary.get('summary', '')}")
    print(f"[monitor] Category: {summary.get('category', 'other')} | "
          f"Priority: {summary.get('priority', 'medium')}")
    if summary.get("action_items"):
        print(f"[monitor] Actions:  {summary['action_items']}")
    email["summary"] = summary

    # ── 4. Meeting gate ───────────────────────────────────────────────────────
    if not _is_meeting_related(summary, email):
        print("[monitor] Not meeting-related — logged.")
        _mark_as_read(msg)
        result = {
            "status":  "non_meeting",
            "email":   email,
            "summary": summary,
        }
        with _session_log_lock:
            _session_log.append(result)
        return result

    # ── 5. Negotiate ────────────────────────────────────────────────────────────────────────────
    # All scheduling decisions — availability, slot selection, counter-proposals,
    # acceptance handling, state persistence — happen inside negotiate() or
    # negotiate_with_human(). This file owns none of that logic.
    #
    # HUMAN NEGOTIATION: route based on whether the sender is an agent or human.
    # is_agent_email() checks for X-AgentSystem: true in the body.
    # Human emails take the negotiate_with_human() path — no slot blocks,
    # natural prose replies, ICS attachment on confirmation.

    # -- 4.5. RAG context query -----------------------------------------------
    # RAG WEEK 12: retrieve prior interaction history with this sender.
    # Attached to the email dict so negotiation agent can use it in prompts.
    prior_context = query_sender_history(
        sender_email = email.get('email', ''),
        account      = _get_account(),
    )
    if prior_context:
        email['prior_context'] = prior_context
        print(f"[rag] Context retrieved for {email.get('email','')}.")

    print("[monitor] Passing to negotiation agent…")
    if is_agent_email(email.get("body", "")):
        neg_result = negotiate(email)
    else:
        print("[monitor] Human sender detected — using human negotiation path.")
        neg_result = negotiate_with_human(email)

    if neg_result is None:
        # negotiate() returns None for emails needing human review (e.g. clarification)
        print("[monitor] Negotiation requires manual review.")
        _mark_as_read(msg)
        result = {
            "status":  "manual_review",
            "email":   email,
            "summary": summary,
        }
        with _session_log_lock:
            _session_log.append(result)
        return result

    action = neg_result.get("action")
    print(f"[monitor] Negotiation action: {action}")

    # ── 6. Send reply ─────────────────────────────────────────────────────────
    reply = neg_result.get("reply")
    # HUMAN NEGOTIATION: pass ICS string if present (non-None only on human confirmed)
    ics   = neg_result.get("ics") if not is_agent_email(email.get("body", "")) else None
    sent  = _send_reply(reply, ics=ics) if reply else False
    if reply:
        print(f"[monitor] Reply sent: {sent}")

    _mark_as_read(msg)

    # ── 7. Log ────────────────────────────────────────────────────────────────
    result = {
        "status":      action,
        "email":       email,
        "summary":     summary,
        "negotiation": neg_result,
        "reply_sent":  sent,
        "escalate":    neg_result.get("escalate", False),
    }

    if action == "confirmed":
        result["calendar"] = neg_result.get("calendar")
        result["slot"]     = neg_result.get("slot")

    with _session_log_lock:
        _session_log.append(result)

    return result


# ── Poll / monitor loop ───────────────────────────────────────────────────────

def poll_inbox() -> list[dict]:
    """
    Check the inbox and process all unread emails.
    Emails are dispatched to worker threads so multiple negotiations
    can proceed concurrently without blocking each other.
    Results are collected once all threads complete.
    """
    if not WIN32_AVAILABLE:
        print("[monitor] win32com not available.")
        return []

    results      = []
    results_lock = threading.Lock()

    try:
        inbox    = get_outlook_inbox(_get_account())
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        unread   = [msg for msg in messages if msg.UnRead]
        print(f"[monitor] Found {len(unread)} unread email(s).")

        if not unread:
            return []

        if len(unread) == 1:
            # Single email - no threading overhead needed
            result = process_email(unread[0])
            if result:
                results.append(result)
            return results

        # Multiple emails - process concurrently
        # Each negotiation writes its own per-thread file so there is
        # no file contention between workers.
        def _worker(msg):
            # COM objects must be initialised per thread on Windows
            pythoncom.CoInitialize()
            try:
                result = process_email(msg)
                if result:
                    with results_lock:
                        results.append(result)
            except Exception as e:
                print(f"[monitor] Worker error: {e}")
            finally:
                pythoncom.CoUninitialize()

        threads = [threading.Thread(target=_worker, args=(msg,), daemon=True)
                   for msg in unread]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=120)  # 2-minute per-email timeout

    except Exception as e:
        print(f"[monitor] Inbox poll error: {e}")

    return results


def _monitor_loop():
    """Background thread: poll on a timer until stop event fires."""
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


# ── Public control API ────────────────────────────────────────────────────────

def start_monitor():
    global _monitor_thread
    if _monitor_thread and _monitor_thread.is_alive():
        print("[monitor] Already running.")
        return

    # Clear session log for this run
    with _session_log_lock:
        _session_log.clear()

    # Purge any stale active negotiations from previous runs.
    # This cancels orphaned tentative calendar entries and resets
    # round counters so repeated tests on the same subject work cleanly.
    from negotiation_agent import purge_active_negotiations
    purge_active_negotiations()

    _stop_event.clear()
    _monitor_thread = threading.Thread(
        target=_monitor_loop, daemon=True, name="EmailMonitor"
    )
    _monitor_thread.start()
    print("[monitor] Monitor started.")


def stop_monitor() -> list[dict]:
    """Stop the monitor and return the full session log for summarisation."""
    global _monitor_thread
    _stop_event.set()
    if _monitor_thread:
        _monitor_thread.join(timeout=5)
    _monitor_thread = None
    with _session_log_lock:
        log_copy = list(_session_log)

    # RAG WEEK 12: index session outcomes into ChromaDB for future context retrieval.
    try:
        indexed = index_session(log_copy, _get_account())
        if indexed:
            print(f"[rag] Session indexed: {indexed} document(s) stored.")
    except Exception as e:
        print(f"[rag] Session index failed (non-critical): {e}")

    return log_copy


def get_session_log() -> list[dict]:
    """Return a snapshot of the current session log without stopping."""
    with _session_log_lock:
        return list(_session_log)


def is_running() -> bool:
    return _monitor_thread is not None and _monitor_thread.is_alive()