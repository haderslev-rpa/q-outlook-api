import os
mail = os.getenv("foldermail")
from q_outlook_api.functionality.mail_api import get_folders


def test_get_folders():

    user_mail = mail

    folders = get_folders(user_mail)

    print("\n✅ FOLDERS:\n")

    for f in folders:
        print(f"{f.get('displayName')} -> {f.get('id')}")


if __name__ == "__main__":
    test_get_folders()