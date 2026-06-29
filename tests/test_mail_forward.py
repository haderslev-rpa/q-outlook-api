import os
mail = os.getenv("foldermail")
from q_outlook_api.functionality.mail_api import forward_mail


def test_forward():

    user = mail
    message_id = "INDSÆT_MAIL_ID"

    forward_mail(
        user,
        message_id,
        ["rujo@haderslev.dk"]
    )

    print("✅ Mail forwarded")


if __name__ == "__main__":
    test_forward()