import os
mail = os.getenv("mail")

from q_outlook_api.functionality.mail_api import get_mails

def test_get_mails():

    mails = get_mails(
        user_mail=mail,
        limit=10,
        apply_rules=True,
        procesnavn="Flytning af filer",
        include_attachments=True,  # ✅ slå til hvis du vil se attachments
        debug=True
    )

    print("\n✅ RESULTAT:\n")

    for m in mails:
        print("📧 Subject:", m.get("subject"))
        print("👤 From:", m.get("from_email"))
        print("📅 Modtaget:", m.get("received_str"))
        print("🤖 Til robot:", m.get("til_robot"))
        print("🤖 Folder:", m.get("folder"))

        # ✅ attachments
        if m.get("attachments"):
            print("📎 Attachments:")
            for a in m.get("attachments"):
                print("   -", a.get("name"), f"({a.get('size')} bytes)")

        print("-" * 50)


if __name__ == "__main__":
    test_get_mails()