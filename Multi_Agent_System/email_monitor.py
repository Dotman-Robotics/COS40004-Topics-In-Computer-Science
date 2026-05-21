import time
import threading
import json
from datetime import datetime
from ollama_client import ask_llm_chat
from utils import extract_first_json, get_outlook_inbox
from negotiation_agent import negotiate
from email_summarizer import summarize_email

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
BLACKLIST_FILE        = "blacklist.json"

_monitor_thread: threading.Thread | None = None
_stop_event     = threading.Event()
_session_log: list[dict] = []
_session_log_lock        = threading.Lock()

_blacklist: set[str] = set()
_blacklist_lock      = threading.Lock()


def _load_blacklist() -> None:
    global _blacklist
    try:
        with open(BLACKLIST_FILE, "r", encoding="utf-8") as f:
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
        with open(BLACKLIST_FILE, "w", encoding="utf-8") as f:
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

MEETING_CATEGORIES = {"meeting_request", "reply_needed"}

def _is_meeting_related(summary: dict, email: dict) -> bool:
    category = summary.get("category", "other")

    # Direct hit from summarizer
    if category in MEETING_CATEGORIES:
        return True

    DETECTOR_SYSTEM = """You are an email classifier. Output JSON only. Never explain.
Determine if the email is related to scheduling a meeting in any way —
this includes fresh requests, counter-proposals, acceptances, and rejections.
Output: {"is_meeting_related": true/false}"""

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

def _send_reply(reply: dict) -> bool:
    """Send a reply dict {to, subject, body} via Outlook COM."""
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


#email processing section

def process_email(msg) -> dict | None:
    
    try:
        email = {
            "from":          msg.SenderName,
            "email":         msg.SenderEmailAddress,
            "subject":       msg.Subject,
            "body":          msg.Body[:1000],  # 1000 chars: enough for agent tags + slot lists
            "received_time": str(msg.ReceivedTime),
        }
    except Exception as e:
        print(f"[monitor] Could not read message: {e}")
        return None

    print(f"\n[monitor] ── New email from {email['from']}: '{email['subject']}'")

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

    summary = summarize_email(email)
    print(f"[monitor] Summary:  {summary.get('summary', '')}")
    print(f"[monitor] Category: {summary.get('category', 'other')} | "
          f"Priority: {summary.get('priority', 'medium')}")
    if summary.get("action_items"):
        print(f"[monitor] Actions:  {summary['action_items']}")
    email["summary"] = summary

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

    print("[monitor] Passing to negotiation agent…")
    neg_result = negotiate(email)

    if neg_result is None:
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

    reply = neg_result.get("reply")
    sent  = _send_reply(reply) if reply else False
    if reply:
        print(f"[monitor] Reply sent: {sent}")

    _mark_as_read(msg)
    result = { #log
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

def poll_inbox() -> list[dict]:
    if not WIN32_AVAILABLE:
        print("[monitor] win32com not available.")
        return []

    results = []
    try:
        inbox    = get_outlook_inbox(MONITOR_ACCOUNT)
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


def _monitor_loop():
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
    with _session_log_lock:
        _session_log.clear()
    _stop_event.clear()
    _monitor_thread = threading.Thread(
        target=_monitor_loop, daemon=True, name="EmailMonitor"
    )
    _monitor_thread.start()
    print("[monitor] Monitor started.")


def stop_monitor() -> list[dict]:

    global _monitor_thread
    _stop_event.set()
    if _monitor_thread:
        _monitor_thread.join(timeout=5)
    _monitor_thread = None
    with _session_log_lock:
        log_copy = list(_session_log)
    return log_copy


def get_session_log() -> list[dict]:
    with _session_log_lock:
        return list(_session_log)


def is_running() -> bool:
    return _monitor_thread is not None and _monitor_thread.is_alive()