from q_outlook_api.functionality.calendar_api import delete_event


def test_delete():

    event_id = "INDSÆT_EVENT_ID"

    delete_event("a-kassesamtaler@haderslev.dk", event_id)

    print("✅ Event slettet")


if __name__ == "__main__":
    test_delete()