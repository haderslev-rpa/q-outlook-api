# q_outlook_api/functionality/mail_rules.py

from q_outlook_api.sharepoint_api.sp_api import get_client


# ✅ FAST KONFIGURATION (konstanter = faste værdier)
SITE_NAME = "Automatisering"
LIST_NAME = "Overblik over mails til robotter"


# -------------------------------------------------
# MATCH SUBJECT (wildcard)
# -------------------------------------------------
def _match_subject(subject, pattern):

    if not subject or not pattern:
        return False

    subject = subject.lower()
    pattern = pattern.lower()

    if "*" in pattern:
        pattern = pattern.replace("*", "")
        return pattern in subject

    return pattern == subject


# -------------------------------------------------
# HENT SHAREPOINT RULES
# -------------------------------------------------
def _get_rules():

    client = get_client()  # funktion (henter client)
    return client.get_mail_rules(SITE_NAME, LIST_NAME)


# -------------------------------------------------
# APPLY SHAREPOINT RULES
# -------------------------------------------------
def apply_sharepoint_rules(
    mails,
    procesnavn,
    mailbox,
    debug=False
):

    rules = _get_rules()

    if not mails or not rules:
        return []

    # filtrer regler
    rules = [r for r in rules if r.get("mailadresse") == mailbox]
    rules = [r for r in rules if r.get("title") == procesnavn]

    if debug:
        print("\n🔎 REGLER:")
        for r in rules:
            print(r)

    if not rules:
        return []

    filtered = []

    for mail in mails:

        subject = mail.get("subject") or ""
        match_found = False

        if debug:
            print("\n📧 MAIL:", subject)

        for rule in rules:

            if _match_subject(subject, rule.get("emne")):
                match_found = True

                if debug:
                    print("✅ MATCH:", rule.get("emne"))

                mail["til_robot"] = True
                mail["procesnavn"] = procesnavn
                filtered.append(mail)
                break

        if not match_found:

            allow_wrong = any(r.get("treat_wrong_subject") for r in rules)

            if allow_wrong:
                if debug:
                    print("⚠️ INGEN MATCH - tilladt")

                mail["til_robot"] = True
                mail["procesnavn"] = procesnavn
                filtered.append(mail)
            else:
                if debug:
                    print("❌ FILTRERET FRA")

    return filtered
