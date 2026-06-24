import requests
import base64
import os

from q_outlook_api.outlook_api import get_client


# -------------------------------------------------
# INTERNAL HELPER (PAGINATION)
# -------------------------------------------------
def _get_all_pages(url, headers):
    """
    Henter ALLE sider via pagination
    """

    all_data = []

    while url:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        data = response.json()

        items = data.get("value", [])
        all_data.extend(items)

        url = data.get("@odata.nextLink")

    return all_data


# -------------------------------------------------
# GET MAILS (ALTID PAGINATION)
# -------------------------------------------------
def get_mails(user_mail, raw=False):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages"

    mails = _get_all_pages(url, headers)

    if raw:
        return mails

    return [
        {
            "id": m.get("id"),
            "subject": m.get("subject"),
            "from": m.get("from", {}).get("emailAddress", {}).get("address"),
            "is_read": m.get("isRead"),
        }
        for m in mails
    ]


# -------------------------------------------------
# SEARCH MAILS (PAGINATION)
# -------------------------------------------------
def search_mails(user_mail, query, raw=False):
    """
    Søg mails med pagination
    """

    client = get_client()
    headers = client.auth.headers()

    headers["ConsistencyLevel"] = "eventual"

    url = (
        f"{client.base}/users/{user_mail}/messages"
        f"?$search=\"{query}\""
    )

    mails = _get_all_pages(url, headers)

    if raw:
        return mails

    return [
        {
            "id": m.get("id"),
            "subject": m.get("subject"),
            "from": m.get("from", {}).get("emailAddress", {}).get("address"),
        }
        for m in mails
    ]


# -------------------------------------------------
# SEND MAIL
# -------------------------------------------------
def send_mail(user_mail, mail):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/sendMail"

    data = {
        "message": {
            "subject": mail["subject"],
            "body": {
                "contentType": "HTML",
                "content": mail["body"]
            },
            "toRecipients": [
                {"emailAddress": {"address": r}}
                for r in mail["to"]
            ]
        },
        "saveToSentItems": True
    }

    r = requests.post(url, headers=headers, json=data, timeout=30)
    r.raise_for_status()

    return True


# -------------------------------------------------
# FORWARD MAIL
# -------------------------------------------------
def forward_mail(user_mail, message_id, to):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}/forward"

    data = {
        "toRecipients": [
            {"emailAddress": {"address": r}}
            for r in to
        ]
    }

    r = requests.post(url, headers=headers, json=data, timeout=30)
    r.raise_for_status()

    return True


# -------------------------------------------------
# REPLY MAIL
# -------------------------------------------------
def reply_mail(user_mail, message_id, body_text):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}/reply"

    data = {
        "message": {
            "body": {
                "contentType": "HTML",
                "content": body_text
            }
        }
    }

    r = requests.post(url, headers=headers, json=data, timeout=30)
    r.raise_for_status()

    return True


# -------------------------------------------------
# MARK AS READ
# -------------------------------------------------
def mark_as_read(user_mail, message_id):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}"

    r = requests.patch(
        url,
        headers=headers,
        json={"isRead": True},
        timeout=30
    )
    r.raise_for_status()

    return True


# -------------------------------------------------
# UPDATE MAIL (labels)
# -------------------------------------------------
def update_mail(user_mail, message_id, updates):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}"

    r = requests.patch(url, headers=headers, json=updates, timeout=30)
    r.raise_for_status()

    return r.json()


# -------------------------------------------------
# GET FOLDERS
# -------------------------------------------------
def get_folders(user_mail):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/mailFolders"

    return _get_all_pages(url, headers)


# -------------------------------------------------
# CREATE FOLDER
# -------------------------------------------------
def create_folder(user_mail, folder_name, parent_folder_id=None):

    client = get_client()
    headers = client.auth.headers()

    if parent_folder_id:
        url = f"{client.base}/users/{user_mail}/mailFolders/{parent_folder_id}/childFolders"
    else:
        url = f"{client.base}/users/{user_mail}/mailFolders"

    data = {"displayName": folder_name}

    r = requests.post(url, headers=headers, json=data, timeout=30)
    r.raise_for_status()

    return r.json()


# -------------------------------------------------
# MOVE MAIL
# -------------------------------------------------
def move_mail(user_mail, message_id, destination_folder_id):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}/move"

    data = {"destinationId": destination_folder_id}

    r = requests.post(url, headers=headers, json=data, timeout=30)
    r.raise_for_status()

    return r.json()


# -------------------------------------------------
# GET ATTACHMENTS
# -------------------------------------------------
def get_attachments(user_mail, message_id):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}/attachments"

    return _get_all_pages(url, headers)


# -------------------------------------------------
# DOWNLOAD ATTACHMENT
# -------------------------------------------------
def download_attachment(
    user_mail,
    message_id,
    attachment_id,
    save_path=None,
    as_bytes=False
):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}/attachments/{attachment_id}"

    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()

    data = r.json()

    file_bytes = base64.b64decode(data["contentBytes"])

    if as_bytes:
        return file_bytes

    if save_path:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)

        with open(save_path, "wb") as f:
            f.write(file_bytes)

        return save_path

    return file_bytes
