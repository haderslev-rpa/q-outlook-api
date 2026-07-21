import requests

from q_outlook_api.functionality.outlook_api import get_client
from q_outlook_api.utils import (
    to_graph_datetime,
    build_attendees,
    format_event,
)


# -------------------------------------------------
# CREATE EVENT
# -------------------------------------------------

def create_event(user_mail, event):
    """
    Opretter en kalenderaftale.
    """

    client = get_client()
    headers = client.auth.headers()

    headers["Prefer"] = (
        'outlook.timezone="W. Europe Standard Time"'
    )

    data = {
        "subject": event["subject"],
        "body": {
            "contentType": event.get(
                "content_type",
                "HTML",
            ),
            "content": event.get(
                "body",
                "",
            ),
        },
        "start": {
            "dateTime": to_graph_datetime(
                event["start"]
            ),
            "timeZone": (
                "Europe/Copenhagen"
            ),
        },
        "end": {
            "dateTime": to_graph_datetime(
                event["end"]
            ),
            "timeZone": (
                "Europe/Copenhagen"
            ),
        },
        "attendees": build_attendees(
            event.get("participants")
        ),
        "isOnlineMeeting": event.get(
            "is_online_meeting",
            True,
        ),
        "onlineMeetingProvider": event.get(
            "online_meeting_provider",
            "teamsForBusiness",
        ),
    }

    url = (
        f"{client.base}/users/"
        f"{user_mail}/events"
    )

    response = requests.post(
        url,
        headers=headers,
        json=data,
        timeout=30,
    )
    response.raise_for_status()

    return format_event(
        response.json(),
        raw=event.get("raw", False),
    )


# -------------------------------------------------
# GET EVENTS
# -------------------------------------------------

def get_events(
    user_mail,
    start_dt,
    end_dt,
    raw=False,
):
    """
    Henter kalenderaftaler i et interval.
    """

    client = get_client()
    headers = client.auth.headers()

    headers["Prefer"] = (
        'outlook.timezone="W. Europe Standard Time"'
    )

    # Konverter dato korrekt.
    start = to_graph_datetime(start_dt)
    end = to_graph_datetime(end_dt)

    # Hent default kalender.
    url_cal = (
        f"{client.base}/users/"
        f"{user_mail}/calendars"
    )

    response = requests.get(
        url_cal,
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()

    calendars = response.json().get(
        "value",
        [],
    )

    calendar = next(
        (
            calendar
            for calendar in calendars
            if calendar.get(
                "isDefaultCalendar"
            )
        ),
        None,
    )

    if not calendar:
        raise ValueError(
            "Standardkalenderen blev ikke fundet."
        )

    # Start-URL.
    url = (
        f"{client.base}/users/{user_mail}"
        f"/calendars/{calendar['id']}"
        f"/calendarView"
        f"?startDateTime={start}"
        f"&endDateTime={end}"
    )

    all_events = []

    # Pagination.
    while url:
        response = requests.get(
            url,
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()

        data = response.json()

        events = data.get("value", [])
        all_events.extend(events)

        url = data.get("@odata.nextLink")

    if raw:
        return all_events

    return [
        format_event(event)
        for event in all_events
    ]


# -------------------------------------------------
# UPDATE EVENT
# -------------------------------------------------

def update_event(
    user_mail,
    event_id,
    event,
):
    """
    Opdaterer en kalenderaftale.
    """

    client = get_client()
    headers = client.auth.headers()

    headers["Prefer"] = (
        'outlook.timezone="W. Europe Standard Time"'
    )

    payload = {}

    if "subject" in event:
        payload["subject"] = event[
            "subject"
        ]

    if "body" in event:
        payload["body"] = {
            "contentType": event.get(
                "content_type",
                "HTML",
            ),
            "content": event["body"],
        }

    if "start" in event:
        payload["start"] = {
            "dateTime": to_graph_datetime(
                event["start"]
            ),
            "timeZone": (
                "Europe/Copenhagen"
            ),
        }

    if "end" in event:
        payload["end"] = {
            "dateTime": to_graph_datetime(
                event["end"]
            ),
            "timeZone": (
                "Europe/Copenhagen"
            ),
        }

    if "participants" in event:
        payload["attendees"] = (
            build_attendees(
                event["participants"]
            )
        )

    if "is_online_meeting" in event:
        payload["isOnlineMeeting"] = (
            event["is_online_meeting"]
        )

    url = (
        f"{client.base}/users/"
        f"{user_mail}/events/{event_id}"
    )

    response = requests.patch(
        url,
        headers=headers,
        json=payload,
        timeout=30,
    )
    response.raise_for_status()

    return format_event(
        response.json(),
        raw=event.get("raw", False),
    )


# -------------------------------------------------
# DELETE EVENT
# -------------------------------------------------

def delete_event(user_mail, event_id):
    """
    Sletter en kalenderaftale.
    """

    client = get_client()
    headers = client.auth.headers()

    url = (
        f"{client.base}/users/"
        f"{user_mail}/events/{event_id}"
    )

    response = requests.delete(
        url,
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()

    return True