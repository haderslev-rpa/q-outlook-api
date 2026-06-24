from q_outlook_api.functionality.mail_api import mark_as_read


def test_mark_read():

    user = "a-kassesamtaler@haderslev.dk"
    message_id = "INDSÆT_MAIL_ID"

    mark_as_read(user, message_id)

    print("✅ Mail markeret som læst")


if __name__ == "__main__":
    test_mark_read()