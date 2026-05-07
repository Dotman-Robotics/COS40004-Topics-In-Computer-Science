import win32com.client as win32
from datetime import timedelta


def get_calendar_events():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)

    events = []

    for item in calendar.Items:
        try:
            start = item.Start
            end = item.End

            try:
                start = start.replace(tzinfo=None)
                end = end.replace(tzinfo=None)
            except:
                pass

            events.append({
                "start": start,
                "end": end,
                "subject": item.Subject
            })
        except:
            continue

    return events


def has_conflict(new_start, new_end, events):
    for event in events:
        try:
            if new_start < event["end"] and new_end > event["start"]:
                return True
        except Exception as e:
            print("[WARN] Skipping event due to mismatch:", e)
    return False


def suggest_alternatives(start, duration_minutes, events):
    suggestions = []

    duration = timedelta(minutes=duration_minutes)

    for i in range(1, 6):
        new_start = start + timedelta(hours=i)
        new_end = new_start + duration

        if not has_conflict(new_start, new_end, events):
            suggestions.append(new_start.strftime("%Y-%m-%d %H:%M"))

    return suggestions