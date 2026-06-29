from datetime import datetime  # klasse (dato/tid)
from zoneinfo import ZoneInfo  # bibliotek (timezone håndtering)


# -------------------------------------------------
# PARSE DATO (tekst → datetime)
# -------------------------------------------------
def parse_datetime(date_str):
    """
    Konverter tekst → datetime
    """
    return datetime.strptime(date_str, "%d-%m-%Y %H:%M")


# -------------------------------------------------
# UTC → DANSK TID (MAILS)
# -------------------------------------------------
def utc_to_danish(dt_str):
    """
    Konverter UTC string → dansk tid
    """

    if not dt_str:
        return None

    dt = datetime.fromisoformat(dt_str.replace("Z", ""))

    # UTC → Europe/Copenhagen
    dt = dt.replace(tzinfo=ZoneInfo("UTC"))
    dk_dt = dt.astimezone(ZoneInfo("Europe/Copenhagen"))

    return dk_dt


# -------------------------------------------------
# DANSK → UTC (FILTER / Graph)
# -------------------------------------------------
def danish_to_utc(dt):
    """
    Konverter dansk tid → UTC (Graph format)
    """

    if isinstance(dt, str):
        dt = datetime.strptime(dt, "%d-%m-%Y")

    dt = dt.replace(tzinfo=ZoneInfo("Europe/Copenhagen"))

    return dt.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")


# -------------------------------------------------
# DATO (START/SLUT DAG) → UTC
# -------------------------------------------------
def to_graph_filter_datetime(dt, end_of_day=False):
    """
    Bruges til Graph $filter (dato intervals)
    """

    if isinstance(dt, str):
        dt = datetime.strptime(dt, "%d-%m-%Y")

    if end_of_day:
        dt = dt.replace(hour=23, minute=59, second=59)
    else:
        dt = dt.replace(hour=0, minute=0, second=0)

    # brug fælles converter
    return danish_to_utc(dt)


# -------------------------------------------------
# GRAPH DATETIME (kalender)
# -------------------------------------------------
def to_graph_datetime(dt):
    """
    Til kalender (uden timezone konvertering)
    """

    if isinstance(dt, str) and "T" in dt:
        return dt

    if isinstance(dt, str):
        dt = parse_datetime(dt)

    return dt.strftime("%Y-%m-%dT%H:%M:%S")


# -------------------------------------------------
# DELTAGERE
# -------------------------------------------------
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


# -------------------------------------------------
# FORMAT EVENT
# -------------------------------------------------
def format_event(event, raw=False):

    if raw:
        return event

    online_meeting = event.get("onlineMeeting") or {}

    return {
        "id": event.get("id"),
        "subject": event.get("subject"),
        "start": event.get("start", {}).get("dateTime"),
        "end": event.get("end", {}).get("dateTime"),
        "timezone": event.get("start", {}).get("timeZone"),
        "teams_link": online_meeting.get("joinUrl")
    }
