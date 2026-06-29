import os
mail = os.getenv("runemail")
from q_outlook_api.functionality.calendar_api import create_event
from q_outlook_api.utils import parse_datetime


def test_post():

    event = {
        "subject": "Test_Emne",
        "body": "Test_Brødtekst",
        "start": "01-04-2026 09:00",
        "end": "01-04-2026 09:30",
        "participants": [
            {"email": mail, "name": "Rune"}
        ]
    }

    result = create_event("a-kassesamtaler@haderslev.dk", event)

    print("\n✅ Oprettet event:")
    print("ID:", result["id"])
    print("Teams link:", result.get("onlineMeeting", {}).get("joinUrl"))


if __name__ == "__main__":
    test_post()
