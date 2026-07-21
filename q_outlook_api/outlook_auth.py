"""
APP LOGIN TIL MICROSOFT GRAPH

Denne fil håndterer token til Outlook/Graph.
Token bliver cachet, så vi ikke henter et nyt token ved hvert API-kald.
"""

import time

import requests
from automation_server_client import AutomationServer, Credential


# -------------------------------------------------
# INIT (miljø)
# -------------------------------------------------

AutomationServer.from_environment()


# -------------------------------------------------
# CREDENTIAL
# -------------------------------------------------

# Samme credential som SharePoint bruges fortsat.
_cred = Credential.get_credential("API_SHAREPOINT")
_cfg = _cred.data


# -------------------------------------------------
# OUTLOOK AUTH
# -------------------------------------------------

class OutlookAuth:
    """
    Håndterer APP-login mod Microsoft Graph.
    """

    def __init__(
        self,
        tenant_id,
        client_id,
        client_secret,
        scope="https://graph.microsoft.com/.default",
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.scope = scope

        self._token = None
        self._token_expiry = 0

    def get_token(self):
        """
        Returnerer et gyldigt Graph-token.

        Et eksisterende token genbruges, indtil der er mindre end
        fem minutter tilbage af tokenets levetid.
        """

        now = time.time()

        if (
            self._token
            and now < self._token_expiry - 300
        ):
            return self._token

        token_url = (
            "https://login.microsoftonline.com/"
            f"{self.tenant_id}/oauth2/v2.0/token"
        )

        response = requests.post(
            token_url,
            data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": self.scope,
                "grant_type": "client_credentials",
            },
            timeout=60,
        )
        response.raise_for_status()

        token_data = response.json()

        self._token = token_data["access_token"]
        self._token_expiry = (
            now
            + int(token_data.get("expires_in", 3600))
        )

        return self._token

    def get_headers(
        self,
        *,
        prefer_plain_text=False,
        immutable_id=True,
    ):
        """
        Bygger standardheaders til Microsoft Graph.

        prefer_plain_text:
            True betyder, at mail-body ønskes som almindelig tekst.

        immutable_id:
            True betyder, at Graph skal returnere stabile mail-id'er.
        """

        headers = {
            "Authorization": (
                f"Bearer {self.get_token()}"
            ),
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

        preferences = []

        if prefer_plain_text:
            preferences.append(
                'outlook.body-content-type="text"'
            )

        if immutable_id:
            preferences.append(
                'IdType="ImmutableId"'
            )

        if preferences:
            headers["Prefer"] = ", ".join(
                preferences
            )

        return headers

    def headers(
        self,
        *,
        prefer_plain_text=False,
        immutable_id=True,
    ):
        """
        Bagudkompatibel metode.

        Eksisterende kode bruger:
            client.auth.headers()

        Ny kode kan bruge:
            client.get_headers()

        Begge dele giver samme resultat.
        """

        return self.get_headers(
            prefer_plain_text=prefer_plain_text,
            immutable_id=immutable_id,
        )