import time
import requests

from automation_server_client import AutomationServer, Credential


# -------------------------------------------------
# INIT (miljø)
# -------------------------------------------------
AutomationServer.from_environment()

# ✅ GENBRUG samme credential som SharePoint
_cred = Credential.get_credential("API_SHAREPOINT")
_cfg = _cred.data


class OutlookAuth:
    """
    Håndterer APP login (token)
    """

    def __init__(self):
        self.tenant_id = _cfg["tenant_id"]
        self.client_id = _cfg["client_id"]
        self.client_secret = _cred.password

        self._token = None
        self._expiry = 0

    def get_token(self):
        if self._token and time.time() < (self._expiry - 300):
            return self._token

        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default"
        }

        r = requests.post(url, data=data, timeout=30)
        r.raise_for_status()

        token = r.json()

        self._token = token["access_token"]
        self._expiry = time.time() + int(token["expires_in"])

        return self._token

    def headers(self):
        return {
            "Authorization": f"Bearer {self.get_token()}",
            "Content-Type": "application/json"
        }