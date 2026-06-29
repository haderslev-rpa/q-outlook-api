import os
mail = os.getenv("foldermail")
from q_outlook_api.functionality.mail_api import get_attachments, download_attachment


def test_attachments():

    user = mail
    message_id = "INDSÆT_MAIL_ID"

    attachments = get_attachments(user, message_id)

    print("\n✅ ATTACHMENTS:\n")

    for a in attachments:
        print(a["name"], "->", a["id"])

    # 🔹 test download (første fil)
    if attachments:
        first = attachments[0]

        path = download_attachment(
            user,
            message_id,
            first["id"],
            save_path=f"/tmp/{first['name']}"
        )

        print("\n✅ Downloadet til:", path)


if __name__ == "__main__":
    test_attachments()