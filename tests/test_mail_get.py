from q_outlook_api.functionality.mail_api import get_mails


def test_get():

    mails = get_mails("a-kassesamtaler@haderslev.dk")

    print("\n✅ MAILS:\n")

    for m in mails[:10]:  # viser kun 10
        print(m)


if __name__ == "__main__":
    test_get()