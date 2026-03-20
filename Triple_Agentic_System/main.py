from planner_agent import planner_agent
from email_agent import email_agent
from calendar_agent import calendar_agent

from planner_agent import planner_agent
from email_agent import email_agent
from calendar_agent import calendar_agent


def run_agent_system(user_input):
    plan = planner_agent(user_input)
    
    print("\n--- PLAN ---")
    print(plan)

    for action in plan.get("actions", []):
        if action["type"] == "email":
            print("\n--- EMAIL AGENT ---")
            print(email_agent(action["input"]))

        elif action["type"] == "calendar":
            print("\n--- CALENDAR AGENT ---")
            print(calendar_agent(action["input"]))


if __name__ == "__main__":
    while True:
        user_input = input("\nEnter request: ")
        run_agent_system(user_input)