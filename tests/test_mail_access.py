"""
TEST: HENT MAILS OG HENT ÉN MAIL VIA ID

Miljøvariabler:
- TEST_MAIL eller foldermail eller mail
- TEST_MAIL_FOLDER er valgfri, standard er inbox
"""

import os
from pprint import pprint

from q_outlook_api.functionality.mail_api import get_mails


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

USER_MAIL = (
    os.getenv("TEST_MAIL")
    or os.getenv("foldermail")
    or os.getenv("mail")
)

FOLDER = os.getenv("TEST_MAIL_FOLDER", "inbox")


# -------------------------------------------------
# HELPER
# -------------------------------------------------

def assert_mail_structure(mail):
    """
    Kontrollerer den fælles mailstruktur.
    """

    required_fields = {
        "id",
        "message_id",
        "internet_message_id",
        "mailbox",
        "subject",
        "body",
        "body_content_type",
        "sender_name",
        "sender_address",
        "received_datetime_utc",
        "received_datetime_danish",
        "categories",
        "has_attachments",
        "attachment_count",
        "attachment_names",
        "attachments",
    }

    missing_fields = required_fields - set(mail.keys())

    assert not missing_fields, (
        f"Disse felter mangler: {sorted(missing_fields)}"
    )

    assert isinstance(
        mail["categories"],
        list,
    ), "categories skal være en liste"

    assert isinstance(
        mail["attachment_names"],
        list,
    ), "attachment_names skal være en liste"

    assert isinstance(
        mail["attachments"],
        list,
    ), "attachments skal være en liste"

    assert (
        mail["attachment_count"]
        == len(mail["attachments"])
    ), "attachment_count stemmer ikke"

    assert mail["attachment_names"] == [
        attachment["name"]
        for attachment in mail["attachments"]
    ], "attachment_names stemmer ikke"


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_mail_access():
    """
    Henter mails og tester standardformatet.
    """

    if not USER_MAIL:
        raise ValueError(
            "Sæt miljøvariablen TEST_MAIL, "
            "foldermail eller mail."
        )

    print("\nHenter mails...\n")

    mails = get_mails(
        user_mail=USER_MAIL,
        folder=FOLDER,
        limit=5,
        include_attachments=True,
    )

    print(f"Antal mails hentet: {len(mails)}")

    if not mails:
        print("Ingen mails blev fundet.")
        return

    for index, mail in enumerate(mails, start=1):
        assert_mail_structure(mail)

        print(f"\n--- MAIL {index} ---")
        print("Emne:", mail["subject"])
        print("Afsender:", mail["sender_address"])
        print(
            "Modtaget dansk tid:",
            mail["received_datetime_danish"],
        )
        print("Kategorier:", mail["categories"])
        print(
            "Vedhæftninger:",
            mail["attachment_names"],
        )

    first_message_id = mails[0]["message_id"]

    print("\nHenter første mail direkte via message_id...\n")

    direct_result = get_mails(
        user_mail=USER_MAIL,
        message_id=first_message_id,
        include_attachments=True,
    )

    assert len(direct_result) == 1, (
        "Hentning via message_id skal give én mail"
    )

    direct_mail = direct_result[0]
    assert_mail_structure(direct_mail)

    assert direct_mail["message_id"] == first_message_id

    print("Direkte hentet mail:")
    pprint(direct_mail)

    print("\nOK: Mailadgang og standardformat virker.")


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_mail_access()