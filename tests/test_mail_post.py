from q_outlook_api.functionality.mail_api import send_mail


def test_post():

    user = "a-kassesamtaler@haderslev.dk"

    mail = {
        "subject": "Test mail fra Python",
        "body": "Dette er en test mail",
        "to": ["rujo@haderslev.dk"]
    }

    send_mail(user, mail)

    print("✅ Mail sendt")


if __name__ == "__main__":
    test_post()
