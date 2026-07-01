import os
mail = os.getenv("mail")

from q_outlook_api.functionality.mail_api import search_mails_filter


def test_search_filter():

    mails = search_mails_filter(
        user_mail=mail,
        date_from="01-06-2026",
        date_to="02-07-2026",
        limit=10,
        apply_rules=False,
        #procesnavn="Oprettelse af sager i Acadre - Arrangementer",
        include_attachments=False,
        #debug=True
    )

    print("\n✅ RESULTAT:\n")

    for m in mails:
        print("📧 Subject:", m.get("subject"))
        print("📅 Modtaget:", m.get("received_str"))
        print("🤖 Til robot:", m.get("til_robot"))
        print("-" * 50)


if __name__ == "__main__":
    test_search_filter()