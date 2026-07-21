"""
FÆLLES OUTLOOK-HJÆLPEFUNKTIONER
"""

from datetime import datetime
from zoneinfo import ZoneInfo


DANISH_TIMEZONE = ZoneInfo(
    "Europe/Copenhagen"
)
UTC_TIMEZONE = ZoneInfo("UTC")


# -------------------------------------------------
# PARSE DATO (tekst -> datetime)
# -------------------------------------------------

def parse_datetime(date_str):
    """
    Konverterer dansk tekst til datetime.

    Understøttede formater:
    - DD-MM-YYYY HH:MM
    - DD-MM-YYYY HH:MM:SS
    """

    if isinstance(date_str, datetime):
        return date_str

    for date_format in (
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(
                date_str,
                date_format,
            )
        except ValueError:
            continue

    raise ValueError(
        f"Ukendt datoformat: {date_str}"
    )


# -------------------------------------------------
# UTC -> DANSK TID (MAILS)
# -------------------------------------------------

def utc_to_danish(dt_str):
    """
    Konverterer Graph UTC-tid til dansk tid.

    Returnerer format:
    DD-MM-YYYY HH:MM:SS
    """

    if not dt_str:
        return None

    if isinstance(dt_str, datetime):
        dt = dt_str
    else:
        normalized = str(dt_str).strip()

        if normalized.endswith("Z"):
            normalized = (
                normalized[:-1]
                + "+00:00"
            )

        dt = datetime.fromisoformat(
            normalized
        )

    if dt.tzinfo is None:
        dt = dt.replace(
            tzinfo=UTC_TIMEZONE
        )

    danish_dt = dt.astimezone(
        DANISH_TIMEZONE
    )

    return danish_dt.strftime(
        "%d-%m-%Y %H:%M:%S"
    )


# -------------------------------------------------
# DANSK -> UTC (FILTER / GRAPH)
# -------------------------------------------------

def danish_to_utc(dt):
    """
    Konverterer dansk datetime til UTC.
    """

    if isinstance(dt, str):
        dt = parse_datetime(dt)

    if dt.tzinfo is None:
        dt = dt.replace(
            tzinfo=DANISH_TIMEZONE
        )

    return dt.astimezone(
        UTC_TIMEZONE
    )


# -------------------------------------------------
# DATO -> GRAPH FILTER DATETIME
# -------------------------------------------------

def to_graph_filter_datetime(
    dt,
    end_of_day=False,
):
    """
    Konverterer dato til Graph-format.

    end_of_day=True sætter tidspunktet til dagens slutning.
    """

    if isinstance(dt, str):
        parsed = None

        for date_format in (
            "%d-%m-%Y %H:%M:%S",
            "%d-%m-%Y %H:%M",
            "%d-%m-%Y",
            "%Y-%m-%d",
        ):
            try:
                parsed = datetime.strptime(
                    dt,
                    date_format,
                )
                break
            except ValueError:
                continue

        if parsed is None:
            raise ValueError(
                f"Ukendt datoformat: {dt}"
            )

        dt = parsed

    if end_of_day:
        dt = dt.replace(
            hour=23,
            minute=59,
            second=59,
            microsecond=999999,
        )

    utc_dt = danish_to_utc(dt)

    return (
        utc_dt
        .isoformat(timespec="seconds")
        .replace("+00:00", "Z")
    )


# -------------------------------------------------
# GRAPH DATETIME (KALENDER)
# -------------------------------------------------

def to_graph_datetime(dt):
    """
    Konverterer datetime til Graph-tekst.

    Returnerer eksempel:
    2026-04-01T09:00:00
    """

    if isinstance(dt, str):
        dt = parse_datetime(dt)

    return dt.strftime(
        "%Y-%m-%dT%H:%M:%S"
    )


# -------------------------------------------------
# DELTAGERE
# -------------------------------------------------

def build_attendees(participants):
    """
    Bygger Graph-deltagere fra adresser.

    Understøtter:
    - almindelig tekst med mailadresse
    - dictionary med "address"
    - dictionary med "email"
    """

    if not participants:
        return []

    attendees = []

    for participant in participants:
        if isinstance(participant, str):
            address = participant
            name = participant
            attendee_type = "required"
        else:
            # Begge feltnavne understøttes,
            # så eksisterende tests fortsat virker.
            address = (
                participant.get("address")
                or participant.get("email")
            )

            if not address:
                raise ValueError(
                    "Deltager mangler "
                    "'address' eller 'email'."
                )

            name = participant.get(
                "name",
                address,
            )

            attendee_type = participant.get(
                "type",
                "required",
            )

        attendees.append(
            {
                "emailAddress": {
                    "address": address,
                    "name": name,
                },
                "type": attendee_type,
            }
        )

    return attendees


# -------------------------------------------------
# FORMAT EVENT
# -------------------------------------------------

def format_event(event, raw=False):
    """
    Formaterer en Graph-kalenderaftale.
    """

    if raw:
        return event

    organizer = (
        (event.get("organizer") or {})
        .get("emailAddress", {})
    )

    return {
        "id": event.get("id"),
        "subject": event.get("subject"),
        "body_preview": event.get(
            "bodyPreview"
        ),
        "body": event.get("body"),
        "start": event.get("start"),
        "end": event.get("end"),
        "location": (
            event.get("location") or {}
        ).get("displayName"),
        "organizer": organizer.get(
            "address"
        ),
        "organizer_name": organizer.get(
            "name"
        ),
        "attendees": event.get(
            "attendees",
            [],
        ),
        "is_online_meeting": event.get(
            "isOnlineMeeting"
        ),
        "online_meeting": event.get(
            "onlineMeeting"
        ),
        # Det gamle Graph-navn bevares,
        # fordi kalender-testen bruger det.
        "onlineMeeting": event.get(
            "onlineMeeting"
        ),
        "online_meeting_url": (
            (event.get("onlineMeeting") or {})
            .get("joinUrl")
        ),
        "web_link": event.get("webLink"),
    }