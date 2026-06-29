import requests


# -------------------------------------------------
# SHAREPOINT CLIENT (klasse (samling af funktioner))
# -------------------------------------------------
class SharePointClient:

    def __init__(self, auth, tenant_name):
        self.auth = auth
        self.tenant_name = tenant_name
        self.base = "https://graph.microsoft.com/v1.0"

    # -------------------------------------------------
    # GET SITE ID (funktion (slår site op))
    # -------------------------------------------------
    def get_site_id(self, site_name):
        url = f"{self.base}/sites/{self.tenant_name}.sharepoint.com:/sites/{site_name}"

        r = requests.get(url, headers=self.auth.headers(), timeout=30)
        r.raise_for_status()

        return r.json()["id"]

    # -------------------------------------------------
    # GET LIST ID
    # -------------------------------------------------
    def get_list_id(self, site_id, list_name):

        url = f"{self.base}/sites/{site_id}/lists"

        r = requests.get(url, headers=self.auth.headers(), timeout=30)
        r.raise_for_status()

        for lst in r.json()["value"]:
            if lst["displayName"] == list_name:
                return lst["id"]

        raise Exception(f"Liste '{list_name}' findes ikke")

    # -------------------------------------------------
    # GET LIST ITEMS RAW
    # -------------------------------------------------
    def get_list_items_raw(self, site_id, list_id):

        url = f"{self.base}/sites/{site_id}/lists/{list_id}/items?expand=fields"

        r = requests.get(url, headers=self.auth.headers(), timeout=30)
        r.raise_for_status()

        return r.json()["value"]

    # -------------------------------------------------
    # GET MAIL RULES
    # -------------------------------------------------
    def get_mail_rules(
        self,
        site_name="Automatisering",
        list_name="Overblik over mails til robotter"
    ):
        """
        Henter SharePoint regler til mail filtrering
        """

        site_id = self.get_site_id(site_name)
        list_id = self.get_list_id(site_id, list_name)

        items = self.get_list_items_raw(site_id, list_id)

        rules = []

        for item in items:

            fields = item.get("fields", {})

            treat_wrong_subject = (
                fields.get("Skalrobottenbehandlemailsmedfork") == "Ja"
            )

            rule = {
                "title": fields.get("Title"),
                "mailadresse": fields.get("Mailadresse"),
                "emne": fields.get("Emne"),
                "treat_wrong_subject": treat_wrong_subject
            }

            rules.append(rule)

        return rules