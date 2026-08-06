import os
from dotenv import load_dotenv

from q_outlook_api.functionality.mail_api import forward_mail


def test_forward():

    load_dotenv()

    user = os.getenv("mail")

    message_id = "AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0AYBBemCttF0CtwoswQOos5AAAgmRyVwAAARIAEAB1bFbMeLN_QoFwm2BYCWr9"

    forward = {
        "subject": "VS: Test videresendt fra Python",
        "body": "Denne mail er videresendt automatisk fra Python.",
        "to": ["rujo@haderslev.dk"],
        "cc": [],
        "bcc": []
    }

    forward_mail(
        user_mail=user,
        message_id=message_id,
        forward=forward
    )

    print("✅ Mail videresendt")


if __name__ == "__main__":
    test_forward()