from q_outlook_api.functionality.calendar_api import get_events
import os
mail = os.getenv("foldermail")

def test_get():

    events = get_events(
        mail,
        "01-04-2026 07:00",
        "01-04-2026 11:00"
    )

    print("\n✅ EVENTS:\n")

    for e in events:
        print(e.get("subject"))
        print(e.get("start"))
        print(e.get("id"))
        print("-" * 40)


if __name__ == "__main__":
    test_get()