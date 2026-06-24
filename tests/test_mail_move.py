from q_outlook_api.functionality.mail_api import move_mail


def test_move():

    user = "a-kassesamtaler@haderslev.dk"

    message_id = "INDSÆT_MAIL_ID"
    folder_id = "INDSÆT_FOLDER_ID"

    result = move_mail(user, message_id, folder_id)

    print("\n✅ Mail flyttet:")
    print(result.get("id"))


if __name__ == "__main__":
    test_move()