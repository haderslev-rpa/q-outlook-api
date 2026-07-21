"""
TEST: HENT VEDHÆFTNINGER SOM BYTES

Miljøvariabler:
- TEST_MAIL eller foldermail
- TEST_MESSAGE_ID er valgfri
- SAVE_ATTACHMENTS_DIR er valgfri

Hvis TEST_MESSAGE_ID ikke er udfyldt, finder testen selv
den første mail med en reel vedhæftning.
"""

import os
from pathlib import Path
from pprint import pprint

from q_outlook_api.functionality.mail_api import (
    get_attachments,
    get_mails,
)


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

USER_MAIL = (
    os.getenv("mail")
    or os.getenv("foldermail")
    or os.getenv("test")
)

MESSAGE_ID = os.getenv("TEST_MESSAGE_ID")
FOLDER = os.getenv("TEST_MAIL_FOLDER", "inbox")
SAVE_FOLDER = os.getenv("SAVE_ATTACHMENTS_DIR")


# -------------------------------------------------
# FIND TESTMAIL
# -------------------------------------------------

def find_message_id_with_attachment():
    """
    Finder den første mail med en reel vedhæftning.
    """

    mails = get_mails(
        user_mail=USER_MAIL,
        folder=FOLDER,
        limit=25,
        include_attachments=True,
    )

    for mail in mails:
        if mail["attachment_count"] > 0:
            print("Testmail fundet:")
            print("Emne:", mail["subject"])
            print(
                "Filnavne:",
                mail["attachment_names"],
            )

            return mail["message_id"]

    return None


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_attachments():
    """
    Henter alle reelle vedhæftninger som bytes.
    """

    if not USER_MAIL:
        raise ValueError(
            "Sæt miljøvariablen TEST_MAIL, "
            "foldermail eller mail."
        )

    message_id = MESSAGE_ID

    if not message_id:
        message_id = find_message_id_with_attachment()

    if not message_id:
        print(
            "Ingen mail med en reel vedhæftning "
            "blev fundet."
        )
        return

    print("\nHenter vedhæftninger som bytes...\n")

    attachments = get_attachments(
        user_mail=USER_MAIL,
        message_id=message_id,
        get_inline=False,
    )

    print(
        f"Antal vedhæftninger hentet: "
        f"{len(attachments)}"
    )

    for index, attachment in enumerate(
        attachments,
        start=1,
    ):
        print(f"\n--- FIL {index} ---")
        print("Navn:", attachment["name"])
        print(
            "Originalt navn:",
            attachment["original_name"],
        )
        print(
            "Attachment-type:",
            attachment["attachment_type"],
        )
        print("Item-type:", attachment["item_type"])
        print(
            "Content-type:",
            attachment["content_type"],
        )
        print("Størrelse:", attachment["size"])
        print("Inline:", attachment["is_inline"])
        print(
            "Antal hentede bytes:",
            len(attachment["content_bytes"]),
        )

        assert attachment["attachment_type"] != (
            "referenceAttachment"
        )

        assert attachment["is_inline"] is False

        assert isinstance(
            attachment["content_bytes"],
            bytes,
        )

        assert len(attachment["content_bytes"]) > 0

    if attachments:
        print("\nFørste attachment-metadata:")
        first_without_bytes = dict(attachments[0])
        first_without_bytes["content_bytes"] = (
            f"<{len(attachments[0]['content_bytes'])} bytes>"
        )
        pprint(first_without_bytes)

    if SAVE_FOLDER:
        save_path = Path(SAVE_FOLDER)
        save_path.mkdir(
            parents=True,
            exist_ok=True,
        )

        print(
            "\nGemmer testfiler i:",
            save_path,
        )

        for attachment in attachments:
            target_file = save_path / attachment["name"]

            target_file.write_bytes(
                attachment["content_bytes"]
            )

            print("Gemt:", target_file)

    print("\nOK: Vedhæftninger blev hentet som bytes.")


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_attachments()
