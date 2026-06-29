import os
mail = os.getenv("foldermail")
from q_outlook_api.functionality.calendar_api import delete_event


def test_delete():

    event_id = "AAMkADNlMGQ2ZWUwLTdjN2QtNDc3Mi05ZDY3LTJkN2M0MzNmZDRkZQBGAAAAAAAstEhSPijZSZvZSYlAcYCyBwCSsc9bV21aQYYxQ7ouvEptAAAA7S4bAABHN7th__DASodgU0Fz6EifAAY5RaGMAAA="

    delete_event("a-kassesamtaler@haderslev.dk", event_id)

    print("✅ Event slettet")


if __name__ == "__main__":
    test_delete()