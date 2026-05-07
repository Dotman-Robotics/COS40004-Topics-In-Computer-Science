import json
from planner_agent import planner_agent
from email_agent import email_agent, send_outlook_email, read_outlook_emails
from calendar_agent import calendar_agent
import datetime

EMAIL_MONITOR_MODE = False

def run_agent_system(user_input, save_json=True):
    plan = planner_agent(user_input)
    
    print("\n--- PLAN ---")
    print(plan)

    output_data = {
        "timestamp": datetime.datetime.now().isoformat(),
        "user_input": user_input,
        "plan": plan,
        "actions": []
    }

    for action in plan.get("actions", []):
        action_type = action.get("type", "").lower()
        action_input = action.get("input", "")

        action_result = {
            "type": action_type,
            "input": action_input
        }

        if action_type == "email":
            print("\n--- EMAIL AGENT ---")
            email_result = email_agent(action_input)
            print(email_result)

            if not email_result.get("error"):
                send_outlook_email(
                    action_input,
                    sender_account="lateralus.lateralus.40004@outlook.com"
                )

            action_result["output"] = email_result

        elif action_type == "calendar":
            print("\n--- CALENDAR AGENT ---")
            calendar_result = calendar_agent(action_input)

            print("LLM Output:")
            print(calendar_result.get("llm_output"))

            print("\nParsed Event:")
            print(calendar_result.get("parsed_event"))

            print("\nICS File:")
            print(calendar_result.get("ics_file"))

            print("\nOutlook Status:")
            print(calendar_result.get("outlook_status"))

            action_result["output"] = calendar_result

        elif action_type == "check_inbox":
            print("\n--- INBOX ---")

            emails = read_outlook_emails()

            for i, email in enumerate(emails, 1):
                print(f"\n--- Email {i} ---")
                print(f"From: {email['from']} ({email['email']})")
                print(f"Subject: {email['subject']}")
                print(f"Received: {email['received_time']}")
                print(f"Body: {email['body']}")
                print("----------------------")

            action_result["output"] = emails

        else:
            print("Unknown action:", action_type)

    output_data["actions"].append(action_result)

    if save_json:
        filename = f"agent_run_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(output_data, f, ensure_ascii=False, indent=4)
        print(f"\nOutput saved to {filename}")

    return output_data

if __name__ == "__main__":
    while True:
        user_input = input("\nEnter request: ")
        run_agent_system(user_input)