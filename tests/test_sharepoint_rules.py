from q_outlook_api.sharepoint_api.sp_api import get_client
import json


def test_sharepoint_rules():

    print("\n🔎 HENTER SHAREPOINT DATA...\n")

    client = get_client()

    # ✅ hent regler (din funktion)
    rules = client.get_mail_rules()

    print("✅ ANTAL REGLER:", len(rules))

    print("\n📦 RAW JSON (som Python ser det):\n")
    print(rules)

    print("\n📦 PÆN JSON (nem at læse):\n")
    print(json.dumps(rules, indent=4, ensure_ascii=False))

    # ✅ vis første element tydeligt
    if rules:
        print("\n🔎 FØRSTE REGEL:\n")
        for key, value in rules[0].items():
            print(f"{key}: {value}")


if __name__ == "__main__":
    test_sharepoint_rules()
