from q_outlook_api.functionality.mail_api import get_mails
import os
mailadresser = os.getenv("mailadresser").split(",")

def test_access():

    users = mailadresser

    for user in users:
        print(f"\n🔎 Tester adgang: {user}")

        try:
            mails = get_mails(user, limit=1)

            print("✅ OK - adgang")
            print("Subject:", mails[0]["subject"] if mails else "Ingen mails")

        except Exception as e:
            print("❌ Ingen adgang")
            print("Fejl:", str(e))


if __name__ == "__main__":
    test_access()
