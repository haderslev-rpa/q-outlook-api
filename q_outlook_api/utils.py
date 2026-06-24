from datetime import datetime  # klasse (dato/tid)


def parse_datetime(date_str):
    """
    Konverter tekst → datetime
    """
    return datetime.strptime(date_str, "%d-%m-%Y %H:%M")


def to_graph_datetime(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def build_attendees(participants):
    attendees = []

    for p in participants or []:
        attendees.append({
            "emailAddress": {
                "address": p["email"],
                "name": p.get("name", "")
            },
            "type": p.get("type", "required")
        })

    return attendees

def format_event(event, raw=False):
    """
    Formatterer event output
    """

    if raw:
        return event

    return {
        "id": event.get("id"),
        "subject": event.get("subject"),
        "start": event.get("start", {}).get("dateTime"),
        "end": event.get("end", {}).get("dateTime"),
        "timezone": event.get("start", {}).get("timeZone"),
        "teams_link": event.get("onlineMeeting", {}).get("joinUrl")
    }
