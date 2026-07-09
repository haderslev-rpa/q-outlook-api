import requests
import base64
import os

from q_outlook_api.functionality.outlook_api import get_client
from q_outlook_api.utils import utc_to_danish
from q_outlook_api.utils import to_graph_filter_datetime
from q_outlook_api.functionality.mail_rules import apply_sharepoint_rules


# -------------------------------------------------
# INTERNAL HELPER (PAGINATION)
# -------------------------------------------------
def _get_all_pages(url, headers):

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
# BUILD RECIPIENTS
# -------------------------------------------------
def _build_recipients(addresses):
    """
    Bygger Graph recipient format.

    addresses:
        Liste med mailadresser.
        Fx ["test@haderslev.dk"]
    """

    if not addresses:
        return []

    return [
        {
            "emailAddress": {
                "address": address
            }
        }
        for address in addresses
        if address
    ]


# -------------------------------------------------
# SEND MAIL
# -------------------------------------------------
def send_mail(user_mail, mail):
    """
    Sender en ny mail.

    user_mail:
        Postkassen der sender mailen.

    mail:
        Dictionary (samling af data) med:
        - subject
        - body
        - to
        - cc
        - bcc
    """

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/sendMail"

    data = {
        "message": {
            "subject": mail.get("subject", ""),
            "body": {
                "contentType": mail.get("content_type", "HTML"),
                "content": mail.get("body", "")
            },
            "toRecipients": _build_recipients(mail.get("to")),
            "ccRecipients": _build_recipients(mail.get("cc")),
            "bccRecipients": _build_recipients(mail.get("bcc"))
        },
        "saveToSentItems": mail.get("save_to_sent_items", True)
    }

    response = requests.post(
        url,
        headers=headers,
        json=data,
        timeout=30
    )
    response.raise_for_status()

    return True

# -------------------------------------------------
# FORWARD MAIL
# -------------------------------------------------
def forward_mail(user_mail, message_id, forward):
    """
    Videresender en mail.

    user_mail:
        Postkassen hvor mailen ligger.

    message_id:
        ID på mailen der skal videresendes.

    forward:
        Dictionary (samling af data) med:
        - to
        - cc
        - bcc
        - subject
        - body
    """

    client = get_client()
    headers = client.auth.headers()

    # -------------------------------------------------
    # 1. Opret forward draft (kladde)
    # -------------------------------------------------
    create_url = f"{client.base}/users/{user_mail}/messages/{message_id}/createForward"

    create_data = {
        "comment": forward.get("body", "")
    }

    create_response = requests.post(
        create_url,
        headers=headers,
        json=create_data,
        timeout=30
    )
    create_response.raise_for_status()

    draft = create_response.json()
    draft_id = draft["id"]

    # -------------------------------------------------
    # 2. Opdater draft med modtagere og evt. subject
    # -------------------------------------------------
    patch_url = f"{client.base}/users/{user_mail}/messages/{draft_id}"

    patch_data = {
        "toRecipients": _build_recipients(forward.get("to")),
        "ccRecipients": _build_recipients(forward.get("cc")),
        "bccRecipients": _build_recipients(forward.get("bcc"))
    }

    # Hvis du sender subject med, overskrives forward-emnet
    if forward.get("subject"):
        patch_data["subject"] = forward.get("subject")

    patch_response = requests.patch(
        patch_url,
        headers=headers,
        json=patch_data,
        timeout=30
    )
    patch_response.raise_for_status()

    # -------------------------------------------------
    # 3. Send draft
    # -------------------------------------------------
    send_url = f"{client.base}/users/{user_mail}/messages/{draft_id}/send"

    send_response = requests.post(
        send_url,
        headers=headers,
        timeout=30
    )
    send_response.raise_for_status()

    return True


# -------------------------------------------------
# GET ATTACHMENTS (metadata - offentlig funktion)
# -------------------------------------------------
def get_attachments(user_mail, message_id):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/messages/{message_id}/attachments"

    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()

    data = r.json().get("value", [])

    return [
        {
            "id": a.get("id"),
            "name": a.get("name"),
            "content_type": a.get("contentType"),
            "size": a.get("size")
        }
        for a in data
    ]

# -------------------------------------------------
# FORMAT MAIL (fælles formattering)
# -------------------------------------------------
# -------------------------------------------------
# FORMAT MAIL (fælles formattering)
# -------------------------------------------------
def _format_mail(m, user_mail, headers, folder, include_attachments=False):
    """
    Formaterer mail til et rent og ensartet output
    """

    # ✅ attachments (kun hvis slået til)
    attachments = []
    if include_attachments and m.get("hasAttachments"):
        attachments = get_attachments(user_mail, m.get("id"))

    # ✅ konverter UTC -> dansk tid (string)
    dt_received = utc_to_danish(m.get("receivedDateTime"))
    dt_sent = utc_to_danish(m.get("sentDateTime"))

    received_danish_str = None
    sent_danish_str = None

    if dt_received:
        received_danish_str = dt_received.strftime("%d-%m-%Y %H:%M:%S")

    if dt_sent:
        sent_danish_str = dt_sent.strftime("%d-%m-%Y %H:%M:%S")

    return {

        # -------------------------
        # CORE
        # -------------------------
        "id": m.get("id"),
        "subject": m.get("subject"),

        "from_email": m.get("from", {}).get("emailAddress", {}).get("address"),
        "from_name": m.get("from", {}).get("emailAddress", {}).get("name"),

        "folder": folder,

        "to": [
            r.get("emailAddress", {}).get("address")
            for r in (m.get("toRecipients") or [])
        ],
        "cc": [
            r.get("emailAddress", {}).get("address")
            for r in (m.get("ccRecipients") or [])
        ],
        "bcc": [
            r.get("emailAddress", {}).get("address")
            for r in (m.get("bccRecipients") or [])
        ],

        "body_text": (m.get("body") or {}).get("content"),

        # -------------------------
        # DATO (REN STRUKTUR)
        # -------------------------
        "received_utc": m.get("receivedDateTime"),
        "received_danish_str": received_danish_str,

        "sent_utc": m.get("sentDateTime"),
        "sent_danish_str": sent_danish_str,

        # -------------------------
        # STATUS
        # -------------------------
        "is_read": m.get("isRead"),
        "has_attachments": m.get("hasAttachments"),

        "attachments": attachments,

        # -------------------------
        # META
        # -------------------------
        "categories": m.get("categories") or [],
        "conversation_id": m.get("conversationId"),

        # -------------------------
        # BUSINESS
        # -------------------------
        "til_robot": m.get("til_robot"),
        "procesnavn": m.get("procesnavn")
    }


# -------------------------------------------------
# GET MAILS
# -------------------------------------------------
def get_mails(
    user_mail,
    folder="inbox",
    raw=False,
    limit=None,
    apply_rules=False,
    procesnavn=None,
    debug=False,
    include_attachments=False
):

    client = get_client()
    headers = client.auth.headers()

    url = f"{client.base}/users/{user_mail}/mailFolders/{folder}/messages"

    all_mails = []

    while url:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        data = response.json()
        mails = data.get("value", [])

        all_mails.extend(mails)

        if limit and len(all_mails) >= limit:
            all_mails = all_mails[:limit]
            break

        url = data.get("@odata.nextLink")

    if apply_rules:

        if not procesnavn:
            raise Exception("procesnavn mangler")

        all_mails = apply_sharepoint_rules(
            mails=all_mails,
            procesnavn=procesnavn,
            mailbox=user_mail,
            debug=debug
        )

    if raw:
        return {
            "formatted": [
                _format_mail(m, user_mail, headers, folder, include_attachments)
                for m in all_mails
            ],
            "raw": all_mails
        }

    return [
        _format_mail(m, user_mail, headers, folder, include_attachments)
        for m in all_mails
    ]


# -------------------------------------------------
# SEARCH MAILS (QUERY)
# -------------------------------------------------
def search_mails_query(
    user_mail,
    query,
    folder="inbox",
    raw=False,
    limit=None,
    apply_rules=False,
    procesnavn=None,
    debug=False,
    include_attachments=False
):

    client = get_client()
    headers = client.auth.headers()
    headers["ConsistencyLevel"] = "eventual"

    url = (
        f"{client.base}/users/{user_mail}/mailFolders/{folder}/messages"
        f"?$search=\"{query}\""
    )

    mails = _get_all_pages(url, headers)

    if limit:
        mails = mails[:limit]

    if apply_rules:
        mails = apply_sharepoint_rules(
            mails=mails,
            procesnavn=procesnavn,
            mailbox=user_mail,
            debug=debug
        )

    if raw:
        return {
            "formatted": [
                _format_mail(m, user_mail, headers, folder, include_attachments)
                for m in mails
            ],
            "raw": mails
        }

    return [
        _format_mail(m, user_mail, headers, folder, include_attachments)
        for m in mails
    ]


# -------------------------------------------------
# SEARCH MAILS (FILTER)
# -------------------------------------------------
def search_mails_filter(
    user_mail,
    subject=None,
    from_email=None,
    date_from=None,
    date_to=None,
    folder="inbox",
    raw=False,
    limit=None,
    apply_rules=False,
    procesnavn=None,
    debug=False,
    include_attachments=False
):

    client = get_client()
    headers = client.auth.headers()

    filter_parts = []

    if date_from:
        filter_parts.append(
            f"receivedDateTime ge {to_graph_filter_datetime(date_from)}"
        )

    if date_to:
        filter_parts.append(
            f"receivedDateTime le {to_graph_filter_datetime(date_to, end_of_day=True)}"
        )

    if from_email:
        filter_parts.append(
            f"from/emailAddress/address eq '{from_email}'"
        )

    query_string = ""
    if filter_parts:
        query_string = "$filter=" + " AND ".join(filter_parts)

    if query_string:
        url = f"{client.base}/users/{user_mail}/mailFolders/{folder}/messages?{query_string}"
    else:
        url = f"{client.base}/users/{user_mail}/mailFolders/{folder}/messages"

    mails = _get_all_pages(url, headers)

    if subject:
        mails = [m for m in mails if subject.lower() in (m.get("subject") or "").lower()]

    if limit:
        mails = mails[:limit]

    if apply_rules:
        mails = apply_sharepoint_rules(
            mails=mails,
            procesnavn=procesnavn,
            mailbox=user_mail,
            debug=debug
        )

    if raw:
        return {
            "formatted": [
                _format_mail(m, user_mail, headers, folder, include_attachments)
                for m in mails
            ],
            "raw": mails
        }

    return [
        _format_mail(m, user_mail, headers, folder, include_attachments)
        for m in mails
    ]


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
