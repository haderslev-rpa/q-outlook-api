import os
from dotenv import load_dotenv

from q_outlook_api.functionality.mail_api import send_mail


def test_post():

    load_dotenv()

    user = os.getenv("mail")
    runemail = os.getenv("runemail")

    mail = {
        "subject": "Test mail fra Python",
        "body": "Dette er en test mail",
        "to": [runemail],
        "cc": [],
        "bcc": []
    }

    send_mail(user, mail)

    print("✅ Mail sendt")


if __name__ == "__main__":
    test_post()