from q_outlook_api.outlook_auth import OutlookAuth


# -------------------------------------------------
# SINGLETON (genbrug)
# -------------------------------------------------
_client = None


def get_client():
    global _client

    if _client:
        return _client

    auth = OutlookAuth()

    _client = OutlookClient(auth)

    return _client


class OutlookClient:
    """
    Client (samler adgang til Graph)
    """

    def __init__(self, auth):
        self.auth = auth
        self.base = "https://graph.microsoft.com/v1.0"
