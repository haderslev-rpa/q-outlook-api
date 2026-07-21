"""
TEST: SØG MAILS MED GRAPH FILTER

Testen bruger mindst ét filter.

Miljøvariabler:
- TEST_MAIL eller mail
- TEST_SUBJECT er valgfri
- TEST_FROM_EMAIL er valgfri
- TEST_DATE_FROM er valgfri
- TEST_DATE_TO er valgfri
"""

import os
from pprint import pprint

from q_outlook_api.functionality.mail_api import (
    search_mails_filter,
)


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

USER_MAIL = (
    os.getenv("TEST_MAIL")
    or os.getenv("mail")
    or os.getenv("foldermail")
)

FOLDER = os.getenv("TEST_MAIL_FOLDER", "inbox")

SUBJECT = os.getenv("TEST_SUBJECT")
FROM_EMAIL = os.getenv("TEST_FROM_EMAIL")
DATE_FROM = os.getenv("TEST_DATE_FROM")
DATE_TO = os.getenv("TEST_DATE_TO")


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_search_filter():
    """
    Søger mails med strukturerede filtre.
    """

    if not USER_MAIL:
        raise ValueError(
            "Sæt miljøvariablen TEST_MAIL, "
            "mail eller foldermail."
        )

    # Hvis ingen testfiltre er sat, bruger vi et
    # meget bredt datofilter.
    date_from = DATE_FROM or "01-01-2020"

    print("\nSøger mails med filter:")
    print("Subject:", SUBJECT)
    print("From:", FROM_EMAIL)
    print("Date from:", date_from)
    print("Date to:", DATE_TO)
    print()

    mails = search_mails_filter(
        user_mail=USER_MAIL,
        subject=SUBJECT,
        from_email=FROM_EMAIL,
        date_from=date_from,
        date_to=DATE_TO,
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
        print(
            "Modtaget:",
            mail["received_datetime_danish"],
        )
        print("Kategorier:", mail["categories"])
        print(
            "Vedhæftninger:",
            mail["attachment_names"],
        )

    if mails:
        print("\nFørste fulde resultat:")
        pprint(mails[0])

    print("\nOK: Filter-søgning virker.")


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_search_filter()