"""
TEST: OPRET KALENDERAFTALE

Miljøvariabler:
- TEST_CALENDAR_MAIL
- foldermail
- TEST_PARTICIPANT_MAIL
- runemail
"""

import os
from pprint import pprint

from q_outlook_api.functionality.calendar_api import create_event


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

USER_MAIL = (
    os.getenv("TEST_CALENDAR_MAIL")
    or os.getenv("foldermail")
    or "a-kassesamtaler@haderslev.dk"
)

PARTICIPANT_MAIL = (
    os.getenv("TEST_PARTICIPANT_MAIL")
    or os.getenv("runemail")
)


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_post():
    """
    Opretter en testaftale.
    """

    if not PARTICIPANT_MAIL:
        raise ValueError(
            "Sæt miljøvariablen "
            "TEST_PARTICIPANT_MAIL eller runemail."
        )

    event = {
        "subject": "Test_Emne",
        "body": "Test_Brødtekst",
        "start": "01-04-2026 09:00",
        "end": "01-04-2026 09:30",
        "participants": [
            {
                "email": PARTICIPANT_MAIL,
                "name": "Rune",
            }
        ],
        "is_online_meeting": True,
    }

    result = create_event(
        user_mail=USER_MAIL,
        event=event,
    )

    assert result.get("id"), (
        "Event mangler id"
    )

    print("\nOprettet event:")
    print("ID:", result["id"])
    print(
        "Teams-link:",
        result.get("online_meeting_url"),
    )

    print("\nFuldt resultat:")
    pprint(result)

    print("\nOK: Kalenderaftale oprettet.")


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_post()