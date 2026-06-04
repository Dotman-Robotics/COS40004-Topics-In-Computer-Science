import json
import datetime
import threading
import traceback
import sys
import os
import argparse


def _parse_args():
    parser = argparse.ArgumentParser(description="Multi-Agent System")
    parser.add_argument(
        "--account", "-a",
        default="zoomertron@outlook.com",
        help="Outlook account email address this instance monitors and sends from"
    )
    parser.add_argument(
        "--port", "-p",
        type=int,
        default=5000,
        help="Port for the Flask GUI server (default: 5000)"
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Run in CLI mode instead of opening the GUI"
    )
    # Parse only known args so pytest / other tools don't break
    args, _ = parser.parse_known_args()
    return args


# Parse args before any imports that read config at module level
_args = _parse_args()

# ── Inject account config before importing modules that use it ────────────────
# email_monitor, calendar_agent, and email_agent all read these from
# environment variables so a single main.py can launch two instances
# on different ports/accounts without file edits.
os.environ["AGENT_ACCOUNT"] = _args.account
os.environ["AGENT_PORT"]    = str(_args.port)

from planner_agent import planner_agent
from email_agent import compose_email, send_outlook_email, read_outlook_emails, display_inbox
from calendar_agent import calendar_agent
from email_monitor import start_monitor, stop_monitor, is_running, poll_inbox
from email_summarizer import summarize_monitor_session, print_session_summary


def run_agent_system(user_input: str, save_json: bool = True) -> dict:
    print("\n--- PLANNING ---")
    plan = planner_agent(user_input)
    print(plan)

    output_data = {
        "timestamp":  datetime.datetime.now().isoformat(),
        "user_input": user_input,
        "plan":       plan,
        "actions":    [],
    }

    account = os.environ.get("AGENT_ACCOUNT", "zoomertron@outlook.com")

    for action in plan.get("actions", []):
        action_type  = (action.get("type") or action.get("name") or "").lower().strip()
        action_input = action.get("input", "")
        action_result = {"type": action_type, "input": action_input, "output": None}

        if action_type == "email":
            print("\n--- EMAIL AGENT ---")
            email_result = compose_email(action_input, context=user_input)
            print(email_result)
            if not email_result.get("error"):
                send_outlook_email(email_result, sender_account=account)
            action_result["output"] = email_result

        elif action_type == "calendar":
            print("\n--- CALENDAR AGENT ---")
            cal_result = calendar_agent(action_input)
            print("Parsed Event:  ", cal_result.get("parsed_event"))
            print("ICS File:      ", cal_result.get("ics_file"))
            print("Outlook Status:", cal_result.get("outlook_status"))
            action_result["output"] = cal_result

        elif action_type == "check_inbox":
            print("\n--- INBOX ---")
            emails = read_outlook_emails(summarize=True)
            display_inbox(emails, batch_summary=True)
            action_result["output"] = emails

        else:
            print(f"[main] Unknown action type: '{action_type}' — skipping.")
            action_result["output"] = {"error": f"Unknown action: {action_type}"}

        output_data["actions"].append(action_result)

    if save_json:
        filename = f"agent_run_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(output_data, f, ensure_ascii=False, indent=4, default=str)
        print(f"\nSaved to {filename}")

    return output_data


# ── GUI launcher ──────────────────────────────────────────────────────────────

def _start_flask(port: int):
    from app import app
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False, threaded=True)


def _wait_for_flask(port: int, timeout: float = 10.0) -> bool:
    import urllib.request
    import time
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{port}/", timeout=1)
            return True
        except Exception:
            time.sleep(0.2)
    return False


def launch_gui(port: int, account: str):
    print(f"[main] Starting instance for account: {account}")
    print(f"[main] Starting Flask backend on port {port}...")

    flask_thread = threading.Thread(
        target=_start_flask, args=(port,), daemon=True, name="FlaskServer"
    )
    flask_thread.start()

    ready = _wait_for_flask(port)
    if not ready:
        print("[main] WARNING: Flask did not start in time — opening anyway.")

    title = f"Agent System — {account}"
    url   = f"http://127.0.0.1:{port}/"

    try:
        import webview
        print("[main] Opening desktop window...")
        window = webview.create_window(
            title    = title,
            url      = url,
            width    = 1200,
            height   = 780,
            min_size = (800, 560),
            resizable= True,
        )
        webview.start()
        print("[main] Window closed — shutting down.")

    except ImportError:
        import webbrowser
        print(f"[main] pywebview not found — opening {url} in browser.")
        webbrowser.open(url)
        try:
            while True:
                threading.Event().wait(60)
        except KeyboardInterrupt:
            print("[main] Shutting down.")


# ── CLI fallback ──────────────────────────────────────────────────────────────

def _cli_loop():
    account = os.environ.get("AGENT_ACCOUNT", "zoomertron@outlook.com")
    print(f"Agent System CLI ready — account: {account}\nType /help for commands.\n")

    while True:
        try:
            user_input = input("Enter request: ").strip()
            if not user_input:
                continue

            if user_input.lower() == "/monitor on":
                if is_running():
                    print("[main] Monitor is already running.")
                else:
                    start_monitor()

            elif user_input.lower() == "/monitor off":
                if not is_running():
                    print("[main] Monitor is not running.")
                else:
                    print("[main] Stopping monitor...")
                    session_log = stop_monitor()
                    if session_log:
                        report   = summarize_monitor_session(session_log)
                        print_session_summary(report)
                        filename = f"session_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                        with open(filename, "w", encoding="utf-8") as f:
                            json.dump(report, f, ensure_ascii=False, indent=4, default=str)
                        print(f"[main] Session summary saved to {filename}")
                    else:
                        print("[main] No emails were processed during this session.")

            elif user_input.lower() == "/monitor status":
                status = "RUNNING" if is_running() else "STOPPED"
                print(f"[main] Monitor status: {status}")
                if is_running():
                    from email_monitor import get_session_log
                    print(f"[main] Emails processed: {len(get_session_log())}")

            elif user_input.lower() == "/monitor poll":
                print("[main] Running manual inbox poll...")
                results = poll_inbox()
                print(f"[main] Processed {len(results)} email(s).")

            elif user_input.lower() == "/help":
                print("""
Commands:
  Any text          → Run agent system on your request
  /monitor on       → Start automatic inbox monitor
  /monitor off      → Stop monitor and show session summary
  /monitor status   → Show whether monitor is running
  /monitor poll     → Manually trigger one inbox check
  /help             → Show this help
  Ctrl+C            → Exit
""")
            else:
                run_agent_system(user_input)

        except KeyboardInterrupt:
            print("\n[main] Shutting down...")
            if is_running():
                session_log = stop_monitor()
                if session_log:
                    report = summarize_monitor_session(session_log)
                    print_session_summary(report)
            print("[main] Goodbye.")
            break
        except Exception:
            traceback.print_exc()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if _args.cli:
        _cli_loop()
    else:
        launch_gui(port=_args.port, account=_args.account)