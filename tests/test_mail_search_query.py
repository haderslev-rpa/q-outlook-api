"""
TEST: SØG MAILS MED GRAPH QUERY

Miljøvariabler:
- TEST_MAIL eller mail
- TEST_MAIL_QUERY er valgfri
- TEST_MAIL_FOLDER er valgfri
"""

import os
from pprint import pprint

from q_outlook_api.functionality.mail_api import (
    search_mails_query,
)


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

USER_MAIL = (
    os.getenv("TEST_MAIL")
    or os.getenv("mail")
    or os.getenv("foldermail")
)

QUERY = os.getenv(
    "TEST_MAIL_QUERY",
    "hasAttachments:true",
)

FOLDER = os.getenv("TEST_MAIL_FOLDER", "inbox")


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_search_query():
    """
    Søger mails med Graph $search.
    """

    if not USER_MAIL:
        raise ValueError(
            "Sæt miljøvariablen TEST_MAIL, "
            "mail eller foldermail."
        )

    print("\nSøger mails med query:")
    print(QUERY)
    print()

    mails = search_mails_query(
        user_mail=USER_MAIL,
        query=QUERY,
        folder=FOLDER,
        limit=5,
        include_attachments=True,
    )

    print(f"Antal fundne mails: {len(mails)}")

    for index, mail in enumerate(mails, start=1):
        assert isinstance(mail["categories"], list)
        assert isinstance(mail["attachment_names"], list)
        assert isinstance(mail["attachments"], list)

        assert (
            mail["attachment_count"]
            == len(mail["attachments"])
        )

        print(f"\n--- RESULTAT {index} ---")
        print("Emne:", mail["subject"])
        print("Afsender:", mail["sender_address"])
        print("Kategorier:", mail["categories"])
        print(
            "Vedhæftninger:",
            mail["attachment_names"],
        )

    if mails:
        print("\nFørste fulde resultat:")
        pprint(mails[0])

    print("\nOK: Query-søgning virker.")


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_search_query()