import requests

from q_outlook_api.functionality.outlook_api import get_client
from q_outlook_api.utils import (
    to_graph_datetime,
    build_attendees,
    format_event
)


# -------------------------------------------------
# CREATE EVENT
# -------------------------------------------------
def create_event(user_mail, event):

    client = get_client()
    headers = client.auth.headers()
    headers["Prefer"] = 'outlook.timezone="W. Europe Standard Time"'

    data = {
        "subject": event["subject"],
        "body": {
            "contentType": "HTML",
            "content": event["body"]
        },
        "start": {
            "dateTime": to_graph_datetime(event["start"]),
            "timeZone": "Europe/Copenhagen"
        },
        "end": {
            "dateTime": to_graph_datetime(event["end"]),
            "timeZone": "Europe/Copenhagen"
        },
        "attendees": build_attendees(event.get("participants")),
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness"
    }

    url = f"{client.base}/users/{user_mail}/events"

    r = requests.post(url, headers=headers, json=data, timeout=30)
    r.raise_for_status()

    return format_event(r.json(), raw=event.get("raw", False))


# -------------------------------------------------
# GET EVENTS
# -------------------------------------------------
def get_events(user_mail, start_dt, end_dt, raw=False):

    client = get_client()
    headers = client.auth.headers()
    headers["Prefer"] = 'outlook.timezone="W. Europe Standard Time"'

    # ✅ konverter dato korrekt
    start = to_graph_datetime(start_dt)
    end = to_graph_datetime(end_dt)

    # ✅ hent default kalender
    url_cal = f"{client.base}/users/{user_mail}/calendars"

    r = requests.get(url_cal, headers=headers, timeout=30)
    r.raise_for_status()

    cal = next(c for c in r.json()["value"] if c["isDefaultCalendar"])

    # ✅ start URL (første side)
    url = (
        f"{client.base}/users/{user_mail}/calendars/{cal['id']}/calendarView"
        f"?startDateTime={start}&endDateTime={end}"
    )

    all_events = []  # liste (samler events)

    # ✅ 🔥 PAGINATION
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()

        data = r.json()

        events = data.get("value", [])
        all_events.extend(events)

        # næste side
        url = data.get("@odata.nextLink")

    # ✅ RAW output
    if raw:
        return all_events

    # ✅ formatted output
    return [format_event(e) for e in all_events]


# -------------------------------------------------
# UPDATE EVENT
# -------------------------------------------------
def update_event(user_mail, event_id, event):

    client = get_client()
    headers = client.auth.headers()
    headers["Prefer"] = 'outlook.timezone="W. Europe Standard Time"'

    payload = {}

    if "subject" in event:
        payload["subject"] = event["subject"]

    if "body" in event:
        payload["body"] = {
            "contentType": "HTML",
            "content": event["body"]
        }

    if "start" in event:
        payload["start"] = {
            "dateTime": to_graph_datetime(event["start"]),
            "timeZone": "Europe/Copenhagen"
        }

    if "end" in event:
        payload["end"] = {
            "dateTime": to_graph_datetime(event["end"]),
            "timeZone": "Europe/Copenhagen"
        }

    if "participants" in event:
        payload["attendees"] = build_attendees(event["participants"])

    url = f"{client.base}/users/{user_mail}/events/{event_id}"

    r = requests.patch(url, headers=headers, json=payload, timeout=30)
    r.raise_for_status()

    return format_event(r.json(), raw=event.get("raw", False))


# -------------------------------------------------
# DELETE EVENT
# -------------------------------------------------
def delete_event(user_mail, event_id):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/events/{event_id}"

    r = requests.delete(url, headers=headers, timeout=30)
    r.raise_for_status()

    return True
