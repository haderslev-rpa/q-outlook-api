import os
mail = os.getenv("mail")

from q_outlook_api.functionality.mail_api import search_mails_query

def test_search_query():

    mails = search_mails_query(
        user_mail=mail,
        query="Test",  # søgetekst
        limit=10,
        apply_rules=True,
        procesnavn="Oprettelse af sager i Acadre - Arrangementer",
        include_attachments=False,
        debug=True
    )

    print("\n✅ RESULTAT:\n")

    for m in mails:
        print("📧 Subject:", m.get("subject"))
        print("👤 From:", m.get("from_email"))
        print("📅 Modtaget:", m.get("received_str"))
        print("-" * 50)


if __name__ == "__main__":
    test_search_query()