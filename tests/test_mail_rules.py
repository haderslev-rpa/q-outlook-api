"""
TEST: HENT MAILS MED SHAREPOINT-REGLER

Miljøvariabler:
- TEST_MAIL
- mail
- foldermail
- TEST_PROCESS_NAME
"""

import os
from pprint import pprint

from q_outlook_api.functionality.mail_api import get_mails


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

USER_MAIL = (
    os.getenv("TEST_MAIL")
    or os.getenv("mail")
    or os.getenv("foldermail")
)

PROCESS_NAME = os.getenv(
    "TEST_PROCESS_NAME",
    "Flytning af filer",
)

FOLDER = os.getenv(
    "TEST_MAIL_FOLDER",
    "inbox",
)


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_get_mails():
    """
    Henter mails og anvender SharePoint-regler.
    """

    if not USER_MAIL:
        raise ValueError(
            "Sæt miljøvariablen TEST_MAIL, "
            "mail eller foldermail."
        )

    mails = get_mails(
        user_mail=USER_MAIL,
        folder=FOLDER,
        limit=10,
        apply_rules=True,
        procesnavn=PROCESS_NAME,
        include_attachments=True,
        debug=True,
    )

    print("\nRESULTAT:\n")

    for mail in mails:
        assert isinstance(
            mail.get("categories"),
            list,
        )

        assert isinstance(
            mail.get("attachment_names"),
            list,
        )

        assert isinstance(
            mail.get("attachments"),
            list,
        )

        print("Subject:", mail.get("subject"))
        print(
            "From:",
            mail.get("sender_address"),
        )
        print(
            "Modtaget:",
            mail.get(
                "received_datetime_danish"
            ),
        )
        print(
            "Til robot:",
            mail.get("til_robot"),
        )
        print(
            "Procesnavn:",
            mail.get("procesnavn"),
        )
        print(
            "Folder:",
            mail.get("folder"),
        )
        print(
            "Kategorier:",
            mail.get("categories"),
        )
        print(
            "Filnavne:",
            mail.get("attachment_names"),
        )

        if mail.get("attachments"):
            print("Attachments:")

            for attachment in mail[
                "attachments"
            ]:
                print(
                    "   -",
                    attachment.get("name"),
                    (
                        f"({attachment.get('size')} "
                        "bytes)"
                    ),
                )

        print("-" * 50)

    if mails:
        print("\nFørste fulde mail:")
        pprint(mails[0])

    print(
        "\nOK: Mailregel-testen blev gennemført."
    )


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_get_mails()