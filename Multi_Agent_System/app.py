import json
import datetime
import os
import threading
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

from planner_agent import planner_agent
from email_agent import email_agent, send_outlook_email, read_outlook_emails, display_inbox
from calendar_agent import calendar_agent
from search_agent import search_agent, draft_outreach_email
from email_monitor import (
    start_monitor, stop_monitor, is_running,
    poll_inbox, get_session_log,
    add_to_blacklist, remove_from_blacklist, get_blacklist,
)
from email_summarizer import summarize_monitor_session
from negotiation_agent import get_all_negotiations, simulate_negotiation

# Resolve GUI folder relative to this file so it works regardless of cwd
_BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
_GUI_FOLDER = os.path.join(_BASE_DIR, "gui")

app = Flask(__name__, static_folder=_GUI_FOLDER, static_url_path="")
CORS(app)

# ── Serve GUI ─────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(_GUI_FOLDER, "index.html")


# ── Agent endpoint ────────────────────────────────────────────────────────────

@app.route("/api/run", methods=["POST"])
def run_agent():
    data       = request.get_json(force=True)
    user_input = (data.get("input") or "").strip()
    if not user_input:
        return jsonify({"error": "No input provided."}), 400

    plan    = planner_agent(user_input)
    actions = []

    for action in plan.get("actions", []):
        action_type  = (action.get("type") or action.get("name") or "").lower().strip()
        action_input = action.get("input", "")
        result_entry = {"type": action_type, "input": action_input, "output": None}

        if action_type == "email":
            email_result = email_agent(action_input, context=user_input)
            if not email_result.get("error"):
                send_outlook_email(email_result, sender_account="zoomertron@outlook.com")
            result_entry["output"] = email_result

        elif action_type == "calendar":
            cal_result = calendar_agent(action_input)
            result_entry["output"] = {
                "parsed_event":  cal_result.get("parsed_event"),
                "ics_file":      cal_result.get("ics_file"),
                "outlook_status": cal_result.get("outlook_status"),
                "availability":  cal_result.get("availability"),
            }

        elif action_type == "check_inbox":
            emails = read_outlook_emails(summarize=True)
            result_entry["output"] = emails

        elif action_type == "search":
            n = int(data.get("n", 5))
            search_result = search_agent(action_input, n=n)
            result_entry["output"] = search_result

        else:
            result_entry["output"] = {"error": f"Unknown action: {action_type}"}

        actions.append(result_entry)

    # Save run log
    log = {
        "timestamp":  datetime.datetime.now().isoformat(),
        "user_input": user_input,
        "plan":       plan,
        "actions":    actions,
    }
    filename = f"agent_run_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=4, default=str)

    return jsonify({"plan": plan, "actions": actions, "log_file": filename})


# ── Monitor endpoints ─────────────────────────────────────────────────────────

@app.route("/api/monitor/start", methods=["POST"])
def monitor_start():
    if is_running():
        return jsonify({"status": "already_running"})
    start_monitor()
    return jsonify({"status": "started"})


@app.route("/api/monitor/stop", methods=["POST"])
def monitor_stop():
    if not is_running():
        return jsonify({"status": "not_running", "summary": None})
    session_log = stop_monitor()
    summary     = summarize_monitor_session(session_log) if session_log else None

    if summary:
        filename = f"session_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=4, default=str)

    return jsonify({"status": "stopped", "summary": summary})


@app.route("/api/monitor/status", methods=["GET"])
def monitor_status():
    running = is_running()
    log     = get_session_log() if running else []
    return jsonify({
        "running":          running,
        "emails_processed": len(log),
    })


@app.route("/api/monitor/poll", methods=["POST"])
def monitor_poll():
    results = poll_inbox()
    return jsonify({"processed": len(results), "results": results})


@app.route("/api/monitor/log", methods=["GET"])
def monitor_log():
    log = get_session_log()
    return jsonify({"log": log, "count": len(log)})


# ── Inbox endpoint ────────────────────────────────────────────────────────────

@app.route("/api/inbox", methods=["GET"])
def inbox():
    max_emails  = int(request.args.get("max", 10))
    unread_only = request.args.get("unread_only", "true").lower() == "true"
    emails      = read_outlook_emails(max_emails=max_emails, unread_only=unread_only, summarize=True)
    return jsonify({"emails": emails, "count": len(emails)})


# ── Search endpoints ──────────────────────────────────────────────────────────

@app.route("/api/search", methods=["POST"])
def search_endpoint():
    data = request.get_json(force=True)
    task = (data.get("task") or "").strip()
    n    = int(data.get("n", 5))
    if not task:
        return jsonify({"error": "No task provided."}), 400
    result = search_agent(task, n=n)
    return jsonify(result)


@app.route("/api/search/draft", methods=["POST"])
def search_draft():
    """Draft an outreach email to a selected provider."""
    data     = request.get_json(force=True)
    provider = data.get("provider")
    task     = (data.get("task")    or "").strip()
    context  = (data.get("context") or "").strip()

    if not provider or not task:
        return jsonify({"error": "provider and task are required."}), 400

    draft = draft_outreach_email(provider, task, extra_context=context)
    return jsonify(draft)


@app.route("/api/search/send", methods=["POST"])
def search_send():
    """Send a drafted outreach email via Outlook."""
    data  = request.get_json(force=True)
    email = data.get("email")
    if not email or not all(k in email for k in ("to", "subject", "body")):
        return jsonify({"error": "email with to/subject/body required."}), 400
    result = send_outlook_email(email, sender_account="zoomertron@outlook.com")
    return jsonify({"sent": result is not None, "email": email})


# ── Blacklist endpoints ───────────────────────────────────────────────────────

@app.route("/api/blacklist", methods=["GET"])
def blacklist_get():
    return jsonify({"blacklist": get_blacklist()})


@app.route("/api/blacklist/add", methods=["POST"])
def blacklist_add():
    data    = request.get_json(force=True)
    address = (data.get("email") or "").strip()
    if not address:
        return jsonify({"error": "No email address provided."}), 400
    added = add_to_blacklist(address)
    return jsonify({"added": added, "address": address.lower(), "blacklist": get_blacklist()})


@app.route("/api/blacklist/remove", methods=["POST"])
def blacklist_remove():
    data    = request.get_json(force=True)
    address = (data.get("email") or "").strip()
    if not address:
        return jsonify({"error": "No email address provided."}), 400
    removed = remove_from_blacklist(address)
    return jsonify({"removed": removed, "address": address.lower(), "blacklist": get_blacklist()})


# ── Negotiation endpoints ─────────────────────────────────────────────────────

@app.route("/api/negotiations", methods=["GET"])
def negotiations_get():
    """Return all negotiation threads (active and closed)."""
    all_neg = get_all_negotiations()
    threads = []
    for thread_id, data in all_neg.items():
        threads.append({
            "thread_id":  thread_id,
            "title":      data.get("title", "Meeting"),
            "status":     data.get("status", "unknown"),
            "rounds":     data.get("rounds", 0),
            "initiator":  data.get("initiator", ""),
            "created_at": data.get("created_at", ""),
            "closed_at":  data.get("closed_at", ""),
            "history":    data.get("history", []),
        })
    threads.sort(key=lambda x: x["created_at"], reverse=True)
    return jsonify({"negotiations": threads, "count": len(threads)})


@app.route("/api/negotiations/simulate", methods=["POST"])
def negotiations_simulate():
    """Run a dry-run negotiation simulation between two in-process agents."""
    data     = request.get_json(force=True)
    title    = data.get("title",    "Project Sync")
    duration = int(data.get("duration", 60))
    slots    = data.get("initial_slots")  # list of {date, time} or None

    transcript = simulate_negotiation(
        title          = title,
        duration       = duration,
        initial_slots  = slots,
        max_sim_rounds = int(data.get("max_rounds", 8)),
    )

    # Serialise for JSON — strip non-serialisable email body depth
    clean = []
    for step in transcript:
        clean.append({
            "round":    step["round"],
            "sender":   step["sender"],
            "receiver": step["receiver"],
            "action":   step["action"],
            "slot":     step.get("slot"),
            "subject":  step["email"].get("subject", ""),
        })

    final = transcript[-1] if transcript else {}
    return jsonify({
        "outcome":    final.get("action", "unknown"),
        "slot":       final.get("slot"),
        "rounds":     len(transcript),
        "transcript": clean,
    })


if __name__ == "__main__":
    print("Starting Agent GUI server at http://localhost:5000")
    app.run(debug=False, port=5000, threaded=True)