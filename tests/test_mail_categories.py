import os

from dotenv import load_dotenv

from q_outlook_api.functionality.mail_api import (
    get_mails,
    update_mail_categories,
)


def test_categories():

    load_dotenv()

    user = os.getenv("mail")

    message_id = (
        "AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0AYBBemCttF0CtwoswQOos5"
        "AAAgmRyVwAAARIAEAB1bFbMeLN_QoFwm2BYCWr9"
    )

    categories = [
        "Test",
        "Point: 300",
        "Nr: 9999",
    ]

    if not user:
        raise ValueError(
            "Miljøvariablen 'mail' mangler."
        )

    # -------------------------------------------------
    # OPDATER KATEGORIER
    # -------------------------------------------------

    update_mail_categories(
        user_mail=user,
        message_id=message_id,
        categories=categories,
    )

    print(
        "✅ Testkategorier tilføjet"
    )

    # -------------------------------------------------
    # HENT MAILEN IGEN OG KONTROLLÉR
    # -------------------------------------------------

    mails = get_mails(
        user_mail=user,
        message_id=message_id,
        include_attachments=False,
        prefer_plain_text=True,
    )

    if not mails:
        raise AssertionError(
            "Mailen kunne ikke hentes efter opdateringen."
        )

    saved_categories = (
        mails[0].get("categories")
        or []
    )

    print(
        "Kategorier i Outlook:",
        saved_categories,
    )

    assert set(saved_categories) == set(
        categories
    ), (
        "Kategorierne blev ikke gemt korrekt. "
        f"Forventet: {categories}. "
        f"Faktisk: {saved_categories}."
    )

    print(
        "✅ Kategorierne blev kontrolleret"
    )


if __name__ == "__main__":
    test_categories()