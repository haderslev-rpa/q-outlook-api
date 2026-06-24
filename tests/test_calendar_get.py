from q_outlook_api.functionality.calendar_api import get_calendar_events


def test_get():

    events = get_calendar_events(
        "a-kassesamtaler@haderslev.dk",
        "2026-01-04T00:00:00",
        "2026-02-04T00:00:00"
    )

    print("\n✅ EVENTS:\n")

    for e in events:
        print(e.get("subject"))
        print(e.get("start"))
        print("-" * 40)


if __name__ == "__main__":
    test_get()