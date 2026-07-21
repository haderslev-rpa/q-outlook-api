"""
FÆLLES MICROSOFT GRAPH CLIENT

Denne fil opretter én fælles Graph-klient, som genbruges af
mail- og kalenderfunktionerne.
"""

from q_outlook_api.outlook_auth import OutlookAuth


# -------------------------------------------------
# SINGLETON (genbrug)
# -------------------------------------------------

_client = None


def get_client():
    """
    Returnerer den fælles OutlookClient.

    Klienten oprettes kun første gang funktionen kaldes.
    """

    global _client

    if _client is None:
        from q_outlook_api.outlook_auth import _cfg, _cred

        auth = OutlookAuth(
            tenant_id=_cfg["tenant_id"],
            client_id=_cfg["client_id"],
            client_secret=_cred.password,
            scope=_cfg.get(
                "scope",
                "https://graph.microsoft.com/.default",
            ),
        )

        _client = OutlookClient(auth)

    return _client


class OutlookClient:
    """
    Samler adgang til Microsoft Graph.
    """

    def __init__(self, auth):
        self.auth = auth
        self.base = "https://graph.microsoft.com/v1.0"

    def get_headers(
        self,
        *,
        prefer_plain_text=False,
        immutable_id=True,
    ):
        """
        Returnerer headers fra auth-objektet.
        """

        return self.auth.get_headers(
            prefer_plain_text=prefer_plain_text,
            immutable_id=immutable_id,
        )