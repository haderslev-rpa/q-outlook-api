"""
MICROSOFT GRAPH MAIL API

Fælles funktioner til:
- hente mails
- søge mails
- hente vedhæftninger
- sende mails
- videresende mails
- flytte mails
- markere mails som læst
- hente mailmapper
- opdatere mailkategorier

VIGTIGT:
- Mail-body kan returneres som almindelig tekst eller HTML.
- Kategorier returneres altid som en liste.
- referenceAttachment returneres aldrig som en vedhæftning.
- Inline-filer returneres kun, når get_inline=True.
- get_mails og søgefunktionerne returnerer metadata, ikke fil-bytes.
- get_attachments returnerer de faktiske filer som bytes i hukommelsen.
"""

import base64
import logging
import os
from pathlib import Path
from urllib.parse import quote

import requests

from q_outlook_api.functionality.mail_rules import apply_sharepoint_rules
from q_outlook_api.functionality.outlook_api import get_client
from q_outlook_api.utils import (
    to_graph_filter_datetime,
    utc_to_danish,
)


# -------------------------------------------------
# KONSTANTER
# -------------------------------------------------

DEFAULT_TIMEOUT = 120

MAIL_SELECT_FIELDS = [
    "id",
    "internetMessageId",
    "conversationId",
    "subject",
    "body",
    "bodyPreview",
    "from",
    "sender",
    "toRecipients",
    "ccRecipients",
    "bccRecipients",
    "replyTo",
    "receivedDateTime",
    "sentDateTime",
    "hasAttachments",
    "categories",
    "importance",
    "isRead",
    "webLink",
]


# -------------------------------------------------
# EGNE FEJLTYPER
# -------------------------------------------------

class MailNotFoundError(Exception):
    """
    Fejltype når en mail ikke længere findes.
    """


class AttachmentDownloadError(Exception):
    """
    Fejltype når en vedhæftning ikke kan hentes.
    """


# -------------------------------------------------
# INTERNAL HELPER (PAGINATION)
# -------------------------------------------------

def _get_all_pages(url, headers, limit=None):
    """
    Henter alle sider fra Microsoft Graph.

    limit:
        Maksimalt antal elementer der skal returneres.
    """

    all_data = []

    while url:
        response = requests.get(
            url,
            headers=headers,
            timeout=DEFAULT_TIMEOUT,
        )
        response.raise_for_status()

        data = response.json()

        items = data.get("value", [])
        all_data.extend(items)

        if limit and len(all_data) >= limit:
            return all_data[:limit]

        url = data.get("@odata.nextLink")

    return all_data


# -------------------------------------------------
# INTERNAL HELPER (URL)
# -------------------------------------------------

def _user_url(user_mail):
    """
    Bygger Graph-URL til en bestemt bruger eller postkasse.
    """

    client = get_client()

    encoded_user_mail = quote(
        str(user_mail),
        safe="",
    )

    return f"{client.base}/users/{encoded_user_mail}"


def _message_url(user_mail, message_id):
    """
    Bygger Graph-URL til en bestemt mail.
    """

    encoded_message_id = quote(
        str(message_id),
        safe="",
    )

    return (
        f"{_user_url(user_mail)}"
        f"/messages/{encoded_message_id}"
    )


# -------------------------------------------------
# INTERNAL HELPER (HEADERS)
# -------------------------------------------------

def _get_headers(
    *,
    prefer_plain_text=True,
    immutable_id=True,
):
    """
    Returnerer Graph-headers.

    prefer_plain_text:
        True giver almindelig tekst.
        False giver Graphs normale body-format.

    immutable_id:
        True giver stabile mail-id'er.
    """

    client = get_client()

    return client.get_headers(
        prefer_plain_text=prefer_plain_text,
        immutable_id=immutable_id,
    )


# -------------------------------------------------
# BUILD RECIPIENTS
# -------------------------------------------------

def _build_recipients(addresses):
    """
    Bygger Graph recipient format.

    addresses:
        Liste med mailadresser.

        Eksempel:
        ["test@haderslev.dk"]

        Der kan også bruges dictionaries:
        [
            {
                "address": "test@haderslev.dk",
                "name": "Test"
            }
        ]
    """

    if not addresses:
        return []

    if isinstance(addresses, str):
        addresses = [addresses]

    recipients = []

    for address in addresses:
        if not address:
            continue

        if isinstance(address, str):
            email_address = {
                "address": address,
            }
        else:
            email_address = {
                "address": address.get("address"),
            }

            if address.get("name"):
                email_address["name"] = address.get("name")

        if not email_address.get("address"):
            continue

        recipients.append(
            {
                "emailAddress": email_address,
            }
        )

    return recipients


# -------------------------------------------------
# FORMAT EMAIL ADDRESS
# -------------------------------------------------

def _format_email_address(value):
    """
    Formaterer én Graph-mailadresse.
    """

    email_address = (
        (value or {}).get("emailAddress")
        or {}
    )

    return {
        "name": email_address.get("name"),
        "address": email_address.get("address"),
    }


def _format_email_addresses(values):
    """
    Formaterer flere Graph-mailadresser.
    """

    return [
        _format_email_address(value)
        for value in (values or [])
    ]


def _get_address_list(values):
    """
    Returnerer kun mailadresserne som en liste.

    Funktionen bevarer de gamle felter:
    - to
    - cc
    - bcc
    """

    return [
        address.get("address")
        for address in _format_email_addresses(values)
        if address.get("address")
    ]


# -------------------------------------------------
# ATTACHMENT TYPE
# -------------------------------------------------

def _get_attachment_type(attachment):
    """
    Finder Microsoft Graph attachment-typen.
    """

    odata_type = (
        attachment.get("@odata.type")
        or ""
    )

    if odata_type.endswith("fileAttachment"):
        return "fileAttachment"

    if odata_type.endswith("itemAttachment"):
        return "itemAttachment"

    if odata_type.endswith("referenceAttachment"):
        return "referenceAttachment"

    return "unknown"


def _get_item_type(attachment):
    """
    Finder typen på et vedhæftet Outlook-item.

    Eksempler:
    - message
    - event
    - contact
    """

    attached_item = attachment.get("item") or {}
    odata_type = attached_item.get("@odata.type") or ""

    if odata_type.endswith("message"):
        return "message"

    if odata_type.endswith("event"):
        return "event"

    if odata_type.endswith("contact"):
        return "contact"

    name = (
        attachment.get("name")
        or ""
    ).lower()

    if name.endswith(".eml") or name.endswith(".msg"):
        return "message"

    return None


# -------------------------------------------------
# ATTACHMENT FILNAVN
# -------------------------------------------------

def _normalize_attachment_name(
    name,
    attachment_type,
    item_type,
):
    """
    Sikrer at en vedhæftning har et brugbart filnavn.

    Vedhæftede mails får filtypen .eml.
    """

    cleaned_name = (name or "").strip()

    if not cleaned_name:
        if (
            attachment_type == "itemAttachment"
            and item_type == "message"
        ):
            cleaned_name = "Vedhæftet mail.eml"
        else:
            cleaned_name = "Vedhæftet fil"

    if (
        attachment_type == "itemAttachment"
        and item_type == "message"
    ):
        lower_name = cleaned_name.lower()

        if lower_name.endswith(".msg"):
            cleaned_name = (
                f"{Path(cleaned_name).stem}.eml"
            )
        elif not lower_name.endswith(".eml"):
            cleaned_name = f"{cleaned_name}.eml"

    return cleaned_name


def _make_attachment_names_unique(attachments):
    """
    Gør dublerede filnavne unikke.

    Eksempel:
    brev.pdf
    brev (1).pdf
    brev (2).pdf
    """

    used_names = set()
    result = []

    for attachment in attachments:
        attachment_copy = dict(attachment)

        original_name = (
            attachment_copy.get("name")
            or "Vedhæftet fil"
        )

        candidate_name = original_name
        sequence = 1

        while candidate_name.casefold() in used_names:
            path = Path(original_name)

            candidate_name = (
                f"{path.stem} ({sequence})"
                f"{path.suffix}"
            )

            sequence += 1

        used_names.add(candidate_name.casefold())

        attachment_copy["original_name"] = (
            attachment_copy.get("original_name")
            or original_name
        )
        attachment_copy["name"] = candidate_name

        result.append(attachment_copy)

    return result


# -------------------------------------------------
# FORMAT ATTACHMENT METADATA
# -------------------------------------------------

def _format_attachment_metadata(
    attachment,
    *,
    get_inline=False,
):
    """
    Formaterer metadata om én vedhæftning.

    Returnerer None hvis vedhæftningen skal filtreres fra.
    """

    attachment_type = _get_attachment_type(
        attachment
    )

    # Reference attachments er cloud-links.
    # De findes allerede i mailens body og skal aldrig
    # returneres som en rigtig vedhæftning.
    if attachment_type == "referenceAttachment":
        return None

    # Ukendte typer returneres ikke.
    if attachment_type == "unknown":
        return None

    is_inline = bool(
        attachment.get("isInline", False)
    )

    # Inline-signaturer og indlejrede billeder
    # returneres kun, når det specifikt ønskes.
    if is_inline and not get_inline:
        return None

    item_type = _get_item_type(attachment)

    normalized_name = _normalize_attachment_name(
        name=attachment.get("name"),
        attachment_type=attachment_type,
        item_type=item_type,
    )

    return {
        "id": attachment.get("id"),
        "name": normalized_name,
        "original_name": attachment.get("name"),
        "attachment_type": attachment_type,
        "item_type": item_type,
        "content_type": attachment.get("contentType"),
        "size": attachment.get("size") or 0,
        "is_inline": is_inline,
        "content_id": attachment.get("contentId"),
        "last_modified_datetime": attachment.get(
            "lastModifiedDateTime"
        ),
    }


# -------------------------------------------------
# GET ATTACHMENT METADATA (intern funktion)
# -------------------------------------------------

def _get_attachment_metadata(
    user_mail,
    message_id,
    *,
    get_inline=False,
):
    """
    Henter oplysninger om vedhæftninger.

    Funktionen henter ikke filernes binære indhold.
    """

    headers = _get_headers(
        prefer_plain_text=True,
    )

    url = (
        f"{_message_url(user_mail, message_id)}"
        f"/attachments"
    )

    raw_attachments = _get_all_pages(
        url=url,
        headers=headers,
    )

    formatted_attachments = []

    for attachment in raw_attachments:
        formatted_attachment = (
            _format_attachment_metadata(
                attachment,
                get_inline=get_inline,
            )
        )

        if formatted_attachment is not None:
            formatted_attachments.append(
                formatted_attachment
            )

    return _make_attachment_names_unique(
        formatted_attachments
    )


# -------------------------------------------------
# SEND MAIL
# -------------------------------------------------

def send_mail(user_mail, mail):
    """
    Sender en ny mail.

    user_mail:
        Postkassen der sender mailen.

    mail:
        Dictionary med:
        - subject
        - body
        - content_type
        - to
        - cc
        - bcc
        - save_to_sent_items
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    url = f"{_user_url(user_mail)}/sendMail"

    data = {
        "message": {
            "subject": mail.get("subject", ""),
            "body": {
                "contentType": mail.get(
                    "content_type",
                    "HTML",
                ),
                "content": mail.get("body", ""),
            },
            "toRecipients": _build_recipients(
                mail.get("to")
            ),
            "ccRecipients": _build_recipients(
                mail.get("cc")
            ),
            "bccRecipients": _build_recipients(
                mail.get("bcc")
            ),
        },
        "saveToSentItems": mail.get(
            "save_to_sent_items",
            True,
        ),
    }

    response = requests.post(
        url,
        headers=headers,
        json=data,
        timeout=DEFAULT_TIMEOUT,
    )
    response.raise_for_status()

    return True

# -------------------------------------------------
# FORWARD MAIL HELPERS
# -------------------------------------------------

logger = logging.getLogger(__name__)


def _validate_forward_mode(
    forward_mode,
):
    """
    Kontrollerer den valgte videresendelsesmetode.

    Tilladte værdier:

    comment:
        En enkel kommentar.

    formatted:
        HTML-formateret tekst indsættes øverst.
    """

    normalized_mode = (
        str(
            forward_mode
            or "comment"
        )
        .strip()
        .casefold()
    )

    valid_modes = {
        "comment",
        "formatted",
    }

    if normalized_mode not in valid_modes:
        raise ValueError(
            "Ukendt forward_mode. "
            "Tilladte værdier er "
            "'comment' og 'formatted'."
        )

    return normalized_mode


def _insert_html_after_body_tag(
    original_html,
    new_html,
):
    """
    Indsætter ny HTML umiddelbart efter
    det eksisterende body-tag.

    Resten af den Graph-genererede HTML
    beholdes uændret.

    Hvis body-tagget ikke findes, placeres
    den nye HTML før det eksisterende indhold.
    """

    original_html = str(
        original_html
        or ""
    )

    new_html = str(
        new_html
        or ""
    )

    body_start = (
        original_html
        .casefold()
        .find("<body")
    )

    if body_start == -1:
        return (
            new_html
            + original_html
        )

    body_tag_end = original_html.find(
        ">",
        body_start,
    )

    if body_tag_end == -1:
        return (
            new_html
            + original_html
        )

    insert_position = (
        body_tag_end + 1
    )

    return (
        original_html[
            :insert_position
        ]
        + new_html
        + original_html[
            insert_position:
        ]
    )


def _delete_forward_draft(
    user_mail,
    draft_id,
):
    """
    Sletter en videresendelseskladde.

    Funktionen bruges kun, hvis en fejl sker,
    før afsendelsen er forsøgt.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    response = requests.delete(
        _message_url(
            user_mail,
            draft_id,
        ),
        headers=headers,
        timeout=DEFAULT_TIMEOUT,
    )

    # 204 betyder, at kladden blev slettet.
    #
    # 404 accepteres også, fordi kladden
    # allerede kan være fjernet.
    if response.status_code not in {
        204,
        404,
    }:
        response.raise_for_status()


def _create_forward_draft(
    user_mail,
    message_id,
    comment=None,
):
    """
    Opretter en videresendelseskladde.

    comment:
        None:
            Opret kladden uden kommentar.

        Tekst:
            Opret kladden med en enkel kommentar.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    create_url = (
        f"{_message_url(user_mail, message_id)}"
        f"/createForward"
    )

    if comment is None:
        create_response = requests.post(
            create_url,
            headers=headers,
            timeout=DEFAULT_TIMEOUT,
        )

    else:
        create_response = requests.post(
            create_url,
            headers=headers,
            json={
                "comment": str(
                    comment
                    or ""
                ),
            },
            timeout=DEFAULT_TIMEOUT,
        )

    if create_response.status_code == 404:
        raise MailNotFoundError(
            "Mailen findes ikke længere i Outlook."
        )

    create_response.raise_for_status()

    draft = create_response.json()

    draft_id = draft.get("id")

    if not draft_id:
        raise ValueError(
            "Microsoft Graph oprettede en kladde, "
            "men svaret mangler draft-id."
        )

    return draft


def _get_forward_draft(
    user_mail,
    draft_id,
):
    """
    Henter videresendelseskladden igen.

    Kladden hentes med HTML-body, så den
    oprindelige videresendelsesformatering
    kan bevares bedst muligt.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    url = (
        f"{_message_url(user_mail, draft_id)}"
        f"?$select=id,subject,body"
    )

    response = requests.get(
        url,
        headers=headers,
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise MailNotFoundError(
            "Videresendelseskladden findes "
            "ikke længere i Outlook."
        )

    response.raise_for_status()

    return response.json()


def _update_forward_draft(
    user_mail,
    draft_id,
    forward,
    body=None,
):
    """
    Opdaterer modtagere, eventuelt emne
    og eventuelt body på kladden.

    body=None:
        Den eksisterende body ændres ikke.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    patch_data = {
        "toRecipients": _build_recipients(
            forward.get("to")
        ),

        "ccRecipients": _build_recipients(
            forward.get("cc")
        ),

        "bccRecipients": _build_recipients(
            forward.get("bcc")
        ),
    }

    if forward.get("subject"):
        patch_data["subject"] = (
            forward.get("subject")
        )

    if body is not None:
        patch_data["body"] = body

    response = requests.patch(
        _message_url(
            user_mail,
            draft_id,
        ),
        headers=headers,
        json=patch_data,
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise MailNotFoundError(
            "Videresendelseskladden findes "
            "ikke længere i Outlook."
        )

    response.raise_for_status()

    return True


def _send_forward_draft(
    user_mail,
    draft_id,
):
    """
    Sender videresendelseskladden.

    Funktionen forsøger kun afsendelsen én gang.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    send_url = (
        f"{_message_url(user_mail, draft_id)}"
        f"/send"
    )

    response = requests.post(
        send_url,
        headers=headers,
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise MailNotFoundError(
            "Videresendelseskladden findes "
            "ikke længere i Outlook."
        )

    response.raise_for_status()

    return True


def _forward_mail_with_comment(
    user_mail,
    message_id,
    forward,
):
    """
    Videresender en mail med en enkel kommentar.

    Graph og Outlook styrer selv den oprindelige
    videresendelsesbody og vedhæftningerne.

    Denne metode giver ikke fuld kontrol
    over HTML-formateringen.
    """

    draft_id = None

    send_attempted = False

    try:
        draft = _create_forward_draft(
            user_mail=user_mail,
            message_id=message_id,
            comment=forward.get(
                "body",
                "",
            ),
        )

        draft_id = draft["id"]

        _update_forward_draft(
            user_mail=user_mail,
            draft_id=draft_id,
            forward=forward,
            body=None,
        )

        # VIGTIGT:
        #
        # Når send_attempted sættes til True,
        # må kladden ikke slettes automatisk.
        #
        # Microsoft Graph kan i sjældne tilfælde
        # have sendt mailen, selv om Python ikke
        # modtager svaret på grund af netværksfejl.

        send_attempted = True

        _send_forward_draft(
            user_mail=user_mail,
            draft_id=draft_id,
        )

        return {
            "success": True,
            "forward_mode": "comment",
            "draft_id": draft_id,
            "send_attempted": True,
        }

    except Exception:

        if (
            draft_id
            and not send_attempted
        ):
            try:
                _delete_forward_draft(
                    user_mail=user_mail,
                    draft_id=draft_id,
                )

                logger.info(
                    "Videresendelseskladde slettet: %s",
                    draft_id,
                )

            except Exception as cleanup_error:
                logger.warning(
                    (
                        "Kunne ikke slette "
                        "videresendelseskladde %s: %s"
                    ),
                    draft_id,
                    cleanup_error,
                )

        raise


def _forward_mail_formatted(
    user_mail,
    message_id,
    forward,
):
    """
    Videresender en mail med HTML-formateret
    tekst øverst.

    Flow:

    1. Graph opretter videresendelseskladden.
    2. Kladdens eksisterende HTML-body hentes.
    3. Den nye HTML indsættes efter body-tagget.
    4. Modtagere og body opdateres.
    5. Kladden sendes.

    Den Graph-genererede videresendelsesbody
    ændres mindst muligt.
    """

    draft_id = None

    send_attempted = False

    try:
        # -------------------------------------------------
        # 1. OPRET FORWARD DRAFT
        # -------------------------------------------------

        draft = _create_forward_draft(
            user_mail=user_mail,
            message_id=message_id,
            comment=None,
        )

        draft_id = draft["id"]


        # -------------------------------------------------
        # 2. HENT KLADDENS HTML-BODY
        # -------------------------------------------------

        full_draft = _get_forward_draft(
            user_mail=user_mail,
            draft_id=draft_id,
        )

        draft_body = (
            full_draft.get("body")
            or {}
        )

        original_content = (
            draft_body.get("content")
            or ""
        )

        original_content_type = (
            draft_body.get(
                "contentType"
            )
            or "HTML"
        ).strip().casefold()

        formatted_body = str(
            forward.get("body")
            or ""
        )


        # -------------------------------------------------
        # 3. BYG SAMLET BODY
        # -------------------------------------------------

        if original_content_type == "html":
            combined_content = (
                _insert_html_after_body_tag(
                    original_html=(
                        original_content
                    ),
                    new_html=(
                        formatted_body
                    ),
                )
            )

        else:
            # Kladden forventes normalt at være HTML.
            #
            # Hvis Graph mod forventning returnerer
            # almindelig tekst, konverteres den
            # oprindelige tekst til sikker HTML.

            from html import escape

            safe_original_content = escape(
                original_content
            ).replace(
                "\n",
                "<br>",
            )

            combined_content = (
                formatted_body
                + (
                    '<div style="'
                    'margin-top: 16px;'
                    'border-top: 1px solid #b7b7b7;'
                    'padding-top: 12px;'
                    '">'
                )
                + safe_original_content
                + "</div>"
            )


        # -------------------------------------------------
        # 4. OPDATER KLADDEN
        # -------------------------------------------------

        _update_forward_draft(
            user_mail=user_mail,
            draft_id=draft_id,
            forward=forward,
            body={
                "contentType": "HTML",
                "content": combined_content,
            },
        )


        # -------------------------------------------------
        # 5. SEND KLADDEN
        # -------------------------------------------------

        # Efter denne linje må kladden ikke
        # slettes automatisk ved fejl.

        send_attempted = True

        _send_forward_draft(
            user_mail=user_mail,
            draft_id=draft_id,
        )

        return {
            "success": True,
            "forward_mode": "formatted",
            "draft_id": draft_id,
            "send_attempted": True,
        }

    except Exception:

        # Kladden slettes kun, hvis send-kaldet
        # endnu ikke er forsøgt.
        #
        # Hvis send er forsøgt, beholdes kladden.
        # Dermed undgår processen at antage,
        # at mailen ikke blev sendt.

        if (
            draft_id
            and not send_attempted
        ):
            try:
                _delete_forward_draft(
                    user_mail=user_mail,
                    draft_id=draft_id,
                )

                logger.info(
                    "Videresendelseskladde slettet: %s",
                    draft_id,
                )

            except Exception as cleanup_error:
                logger.warning(
                    (
                        "Kunne ikke slette "
                        "videresendelseskladde %s: %s"
                    ),
                    draft_id,
                    cleanup_error,
                )

        raise


# -------------------------------------------------
# FORWARD MAIL
# -------------------------------------------------

def forward_mail(
    user_mail,
    message_id,
    forward,
):
    """
    Videresender en eksisterende mail.

    forward skal være en dictionary med:

    {
        "to": [],
        "cc": [],
        "bcc": [],
        "subject": "...",
        "body": "...",
        "forward_mode": "comment"
    }

    forward_mode:

    comment:
        Bruger en enkel kommentar.

        Dette er standardværdien og bevarer
        eksisterende processers adfærd.

    formatted:
        Indsætter HTML-formateret tekst øverst
        i den Graph-genererede videresendelse.

    Hvis forward_mode ikke sendes med,
    bruges comment.
    """

    if not isinstance(
        forward,
        dict,
    ):
        raise TypeError(
            "forward skal være en dictionary."
        )

    forward_mode = _validate_forward_mode(
        forward.get(
            "forward_mode",
            "comment",
        )
    )

    if forward_mode == "comment":
        return _forward_mail_with_comment(
            user_mail=user_mail,
            message_id=message_id,
            forward=forward,
        )

    return _forward_mail_formatted(
        user_mail=user_mail,
        message_id=message_id,
        forward=forward,
    )

# -------------------------------------------------
# DOWNLOAD RAW ATTACHMENT
# -------------------------------------------------

def _download_attachment_value(
    user_mail,
    message_id,
    attachment_id,
):
    """
    Henter rå bytes fra Graphs $value-endpoint.

    Funktionen bruges blandt andet til:
    - almindelige filer uden contentBytes
    - vedhæftede mails som MIME/EML
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    binary_headers = dict(headers)
    binary_headers["Accept"] = "*/*"

    encoded_attachment_id = quote(
        str(attachment_id),
        safe="",
    )

    url = (
        f"{_message_url(user_mail, message_id)}"
        f"/attachments/{encoded_attachment_id}"
        f"/$value"
    )

    response = requests.get(
        url,
        headers=binary_headers,
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise AttachmentDownloadError(
            "Vedhæftningen findes ikke længere."
        )

    try:
        response.raise_for_status()
    except requests.HTTPError as error:
        raise AttachmentDownloadError(
            "Vedhæftningen kunne ikke hentes. "
            f"HTTP-status: {response.status_code}"
        ) from error

    return {
        "content_bytes": response.content,
        "content_type": response.headers.get(
            "Content-Type"
        ),
    }


# -------------------------------------------------
# GET ATTACHMENTS (indhold - offentlig funktion)
# -------------------------------------------------

def get_attachments(
    user_mail,
    message_id,
    get_inline=False,
):
    """
    Henter alle reelle vedhæftninger som bytes.

    Returnerer:
    - almindelige filer som bytes
    - vedhæftede mails som EML-bytes

    Returnerer aldrig:
    - referenceAttachment
    - inline-filer når get_inline=False

    Den vedhæftede mails body og egne vedhæftninger
    bliver ikke læst eller behandlet.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    url = (
        f"{_message_url(user_mail, message_id)}"
        f"/attachments"
    )

    raw_attachments = _get_all_pages(
        url=url,
        headers=headers,
    )

    results = []

    for attachment in raw_attachments:
        metadata = _format_attachment_metadata(
            attachment,
            get_inline=get_inline,
        )

        if metadata is None:
            continue

        attachment_type = metadata[
            "attachment_type"
        ]

        content_bytes = None

        # -------------------------------------------------
        # ALMINDELIG FIL
        # -------------------------------------------------

        if attachment_type == "fileAttachment":
            encoded_content = attachment.get(
                "contentBytes"
            )

            if encoded_content:
                content_bytes = base64.b64decode(
                    encoded_content
                )
            else:
                downloaded = _download_attachment_value(
                    user_mail=user_mail,
                    message_id=message_id,
                    attachment_id=metadata["id"],
                )

                content_bytes = downloaded[
                    "content_bytes"
                ]

                if downloaded.get("content_type"):
                    metadata["content_type"] = (
                        downloaded["content_type"]
                    )

        # -------------------------------------------------
        # VEDHÆFTET OUTLOOK-ITEM
        # -------------------------------------------------

        elif attachment_type == "itemAttachment":
            try:
                downloaded = _download_attachment_value(
                    user_mail=user_mail,
                    message_id=message_id,
                    attachment_id=metadata["id"],
                )
            except AttachmentDownloadError:
                # Et ikke-understøttet Outlook-item
                # skal ikke stoppe de øvrige filer.
                continue

            content_bytes = downloaded[
                "content_bytes"
            ]

            downloaded_content_type = (
                downloaded.get("content_type")
                or ""
            )

            # Graph returnerer en vedhæftet mail som MIME.
            if (
                "message/rfc822"
                in downloaded_content_type.lower()
            ):
                metadata["item_type"] = "message"
                metadata["content_type"] = (
                    "message/rfc822"
                )

                metadata["name"] = (
                    _normalize_attachment_name(
                        name=metadata.get("name"),
                        attachment_type=(
                            "itemAttachment"
                        ),
                        item_type="message",
                    )
                )
            elif downloaded_content_type:
                metadata["content_type"] = (
                    downloaded_content_type
                )

        else:
            continue

        if not isinstance(content_bytes, bytes):
            continue

        metadata["content_bytes"] = content_bytes

        results.append(metadata)

    return _make_attachment_names_unique(results)


# -------------------------------------------------
# FORMAT MAIL (fælles formattering)
# -------------------------------------------------

def _format_mail(
    m,
    user_mail,
    folder,
    include_attachments=False,
    get_inline=False,
):
    """
    Formaterer mail til et rent og ensartet output.

    Funktionen bevarer eksisterende feltnavne og
    tilføjer de nye standardfelter.
    """

    # -------------------------------------------------
    # ATTACHMENT METADATA
    # -------------------------------------------------

    attachment_metadata = []

    # Filnavne skal altid kunne komme med fra get_mails
    # og søgefunktionerne.
    if m.get("hasAttachments") and m.get("id"):
        attachment_metadata = (
            _get_attachment_metadata(
                user_mail=user_mail,
                message_id=m.get("id"),
                get_inline=get_inline,
            )
        )

    # Det fulde metadata-output vises kun, når
    # include_attachments=True.
    if include_attachments:
        attachments = attachment_metadata
    else:
        attachments = []

    attachment_names = [
        attachment.get("name")
        for attachment in attachment_metadata
        if attachment.get("name")
    ]

    attachment_count = len(
        attachment_metadata
    )

    # -------------------------------------------------
    # DATO
    # -------------------------------------------------

    # utc_to_danish returnerer allerede en tekststreng.
    received_danish_str = utc_to_danish(
        m.get("receivedDateTime")
    )

    sent_danish_str = utc_to_danish(
        m.get("sentDateTime")
    )

    # -------------------------------------------------
    # AFSENDER
    # -------------------------------------------------

    from_data = _format_email_address(
        m.get("from")
    )

    sender_data = _format_email_address(
        m.get("sender")
    )

    # Graphs "from" bruges først.
    # "sender" bruges som fallback.
    selected_sender = from_data

    if not selected_sender.get("address"):
        selected_sender = sender_data

    # -------------------------------------------------
    # BODY
    # -------------------------------------------------

    body = m.get("body") or {}

    body_content = body.get("content") or ""
    body_content_type = (
        body.get("contentType")
        or ""
    ).lower()

    # -------------------------------------------------
    # KATEGORIER
    # -------------------------------------------------

    categories = m.get("categories") or []

    if not isinstance(categories, list):
        categories = [str(categories)]

    # -------------------------------------------------
    # RETURNER ENSARTET MAIL
    # -------------------------------------------------

    return {
        # -------------------------
        # CORE
        # -------------------------
        "id": m.get("id"),
        "message_id": m.get("id"),
        "internet_message_id": m.get(
            "internetMessageId"
        ),
        "conversation_id": m.get(
            "conversationId"
        ),
        "mailbox": user_mail,
        "subject": m.get("subject") or "",
        "folder": folder,

        # -------------------------
        # AFSENDER
        # -------------------------
        "sender_name": selected_sender.get(
            "name"
        ),
        "sender_address": selected_sender.get(
            "address"
        ),
        "sender": selected_sender,
        "from": from_data,

        # Gamle felter bevares.
        "from_email": selected_sender.get(
            "address"
        ),
        "from_name": selected_sender.get(
            "name"
        ),

        # -------------------------
        # MODTAGERE
        # -------------------------
        "to": _get_address_list(
            m.get("toRecipients")
        ),
        "cc": _get_address_list(
            m.get("ccRecipients")
        ),
        "bcc": _get_address_list(
            m.get("bccRecipients")
        ),

        "to_recipients": _format_email_addresses(
            m.get("toRecipients")
        ),
        "cc_recipients": _format_email_addresses(
            m.get("ccRecipients")
        ),
        "bcc_recipients": _format_email_addresses(
            m.get("bccRecipients")
        ),
        "reply_to": _format_email_addresses(
            m.get("replyTo")
        ),

        # -------------------------
        # BODY
        # -------------------------
        "body": body_content,
        "body_text": body_content,
        "body_content_type": body_content_type,
        "body_preview": m.get(
            "bodyPreview"
        ) or "",

        # -------------------------
        # DATO
        # -------------------------
        "received_datetime_utc": m.get(
            "receivedDateTime"
        ),
        "received_datetime_danish": (
            received_danish_str
        ),

        "sent_datetime_utc": m.get(
            "sentDateTime"
        ),
        "sent_datetime_danish": (
            sent_danish_str
        ),

        # Gamle datofelter bevares.
        "received_utc": m.get(
            "receivedDateTime"
        ),
        "received_danish_str": (
            received_danish_str
        ),
        "sent_utc": m.get(
            "sentDateTime"
        ),
        "sent_danish_str": sent_danish_str,

        # -------------------------
        # STATUS
        # -------------------------
        "is_read": m.get("isRead"),
        "importance": m.get("importance"),

        # has_attachments viser reelle attachments
        # efter filtrering.
        "has_attachments": (
            attachment_count > 0
        ),
        "attachment_count": attachment_count,
        "attachment_names": attachment_names,
        "attachments": attachments,

        # -------------------------
        # META
        # -------------------------
        "categories": categories,
        "web_link": m.get("webLink"),

        # -------------------------
        # BUSINESS
        # -------------------------
        "til_robot": m.get("til_robot"),
        "procesnavn": m.get("procesnavn"),
    }


# -------------------------------------------------
# FORMAT FLERE MAILS
# -------------------------------------------------

def _format_mails(
    mails,
    user_mail,
    folder,
    include_attachments=False,
    get_inline=False,
):
    """
    Formaterer flere mails med samme struktur.
    """

    return [
        _format_mail(
            m=mail,
            user_mail=user_mail,
            folder=folder,
            include_attachments=(
                include_attachments
            ),
            get_inline=get_inline,
        )
        for mail in mails
    ]


# -------------------------------------------------
# GET MAILS
# -------------------------------------------------

def get_mails(
    user_mail,
    folder="inbox",
    message_id=None,
    raw=False,
    limit=None,
    apply_rules=False,
    procesnavn=None,
    debug=False,
    include_attachments=False,
    get_inline=False,
    prefer_plain_text=True,
):
    """
    Henter mails eller én konkret mail.

    message_id:
        Hvis message_id er udfyldt, hentes kun
        den konkrete mail.

    prefer_plain_text:
        True:
            Body returneres som almindelig tekst.

        False:
            Body returneres i Graphs normale format,
            som oftest er HTML.

    include_attachments:
        True:
            attachments indeholder metadata.

        False:
            attachments er en tom liste.

        attachment_names og attachment_count returneres
        i begge tilfælde.

    Funktionen returnerer altid en liste, medmindre
    raw=True. Ved raw=True returneres både formaterede
    og rå Graph-data.
    """

    headers = _get_headers(
        prefer_plain_text=prefer_plain_text,
    )

    selected_fields = ",".join(
        MAIL_SELECT_FIELDS
    )

    # -------------------------------------------------
    # HENT ÉN KONKRET MAIL
    # -------------------------------------------------

    if message_id:
        url = (
            f"{_message_url(user_mail, message_id)}"
            f"?$select={selected_fields}"
        )

        response = requests.get(
            url,
            headers=headers,
            timeout=DEFAULT_TIMEOUT,
        )

        if response.status_code == 404:
            raise MailNotFoundError(
                "Mailen findes ikke længere i Outlook."
            )

        response.raise_for_status()

        all_mails = [response.json()]

    # -------------------------------------------------
    # HENT MAILS FRA MAPPE
    # -------------------------------------------------

    else:
        encoded_folder = quote(
            str(folder),
            safe="",
        )

        top = limit if limit else 100

        url = (
            f"{_user_url(user_mail)}"
            f"/mailFolders/{encoded_folder}"
            f"/messages"
            f"?$select={selected_fields}"
            f"&$top={top}"
            f"&$orderby=receivedDateTime desc"
        )

        all_mails = _get_all_pages(
            url=url,
            headers=headers,
            limit=limit,
        )

    # -------------------------------------------------
    # SHAREPOINT MAILREGLER
    # -------------------------------------------------

    if apply_rules:
        if not procesnavn:
            raise ValueError(
                "procesnavn mangler"
            )

        all_mails = apply_sharepoint_rules(
            mails=all_mails,
            procesnavn=procesnavn,
            mailbox=user_mail,
            debug=debug,
        )

    formatted_mails = _format_mails(
        mails=all_mails,
        user_mail=user_mail,
        folder=folder,
        include_attachments=include_attachments,
        get_inline=get_inline,
    )

    if raw:
        return {
            "formatted": formatted_mails,
            "raw": all_mails,
        }

    return formatted_mails


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
    include_attachments=False,
    get_inline=False,
    prefer_plain_text=True,
):
    """
    Søger mails med Microsoft Graph $search.

    Eksempler:
    subject:test
    from:test@haderslev.dk
    hasAttachments:true
    """

    headers = _get_headers(
        prefer_plain_text=prefer_plain_text,
    )

    headers["ConsistencyLevel"] = "eventual"

    encoded_folder = quote(
        str(folder),
        safe="",
    )

    encoded_query = quote(
        f'"{query}"',
        safe="",
    )

    selected_fields = ",".join(
        MAIL_SELECT_FIELDS
    )

    top = limit if limit else 100

    url = (
        f"{_user_url(user_mail)}"
        f"/mailFolders/{encoded_folder}"
        f"/messages"
        f"?$search={encoded_query}"
        f"&$select={selected_fields}"
        f"&$top={top}"
    )

    mails = _get_all_pages(
        url=url,
        headers=headers,
        limit=limit,
    )

    if apply_rules:
        if not procesnavn:
            raise ValueError(
                "procesnavn mangler"
            )

        mails = apply_sharepoint_rules(
            mails=mails,
            procesnavn=procesnavn,
            mailbox=user_mail,
            debug=debug,
        )

    formatted_mails = _format_mails(
        mails=mails,
        user_mail=user_mail,
        folder=folder,
        include_attachments=include_attachments,
        get_inline=get_inline,
    )

    if raw:
        return {
            "formatted": formatted_mails,
            "raw": mails,
        }

    return formatted_mails


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
    include_attachments=False,
    get_inline=False,
    prefer_plain_text=True,
):
    """
    Søger mails med Microsoft Graph $filter.

    subject:
        Filtreres lokalt efter Graph-kaldet, så
        eksisterende adfærd bevares.
    """

    headers = _get_headers(
        prefer_plain_text=prefer_plain_text,
    )

    encoded_folder = quote(
        str(folder),
        safe="",
    )

    selected_fields = ",".join(
        MAIL_SELECT_FIELDS
    )

    filter_parts = []

    if date_from:
        filter_parts.append(
            "receivedDateTime ge "
            f"{to_graph_filter_datetime(date_from)}"
        )

    if date_to:
        filter_parts.append(
            "receivedDateTime le "
            f"{to_graph_filter_datetime(date_to, end_of_day=True)}"
        )

    if from_email:
        safe_from_email = str(
            from_email
        ).replace("'", "''")

        filter_parts.append(
            "from/emailAddress/address eq "
            f"'{safe_from_email}'"
        )

    top = limit if limit else 100

    url = (
        f"{_user_url(user_mail)}"
        f"/mailFolders/{encoded_folder}"
        f"/messages"
        f"?$select={selected_fields}"
        f"&$top={top}"
        f"&$orderby=receivedDateTime desc"
    )

    if filter_parts:
        filter_expression = " and ".join(
            filter_parts
        )

        encoded_filter = quote(
            filter_expression,
            safe="'():@.-",
        )

        url = f"{url}&$filter={encoded_filter}"

    mails = _get_all_pages(
        url=url,
        headers=headers,
        limit=limit,
    )

    # Bevarer den eksisterende lokale emnesøgning.
    if subject:
        subject_lower = str(subject).lower()

        mails = [
            mail
            for mail in mails
            if subject_lower
            in (
                mail.get("subject")
                or ""
            ).lower()
        ]

    if limit:
        mails = mails[:limit]

    if apply_rules:
        if not procesnavn:
            raise ValueError(
                "procesnavn mangler"
            )

        mails = apply_sharepoint_rules(
            mails=mails,
            procesnavn=procesnavn,
            mailbox=user_mail,
            debug=debug,
        )

    formatted_mails = _format_mails(
        mails=mails,
        user_mail=user_mail,
        folder=folder,
        include_attachments=include_attachments,
        get_inline=get_inline,
    )

    if raw:
        return {
            "formatted": formatted_mails,
            "raw": mails,
        }

    return formatted_mails


# -------------------------------------------------
# DOWNLOAD ATTACHMENT
# -------------------------------------------------

def download_attachment(
    user_mail,
    message_id,
    attachment_id,
    save_path=None,
    as_bytes=False,
):
    """
    Henter én bestemt vedhæftning.

    Funktionen understøtter:
    - almindelige filer
    - vedhæftede Outlook-items via $value

    Nye processer bør normalt bruge get_attachments,
    fordi den returnerer alle reelle vedhæftninger.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    encoded_attachment_id = quote(
        str(attachment_id),
        safe="",
    )

    url = (
        f"{_message_url(user_mail, message_id)}"
        f"/attachments/{encoded_attachment_id}"
    )

    response = requests.get(
        url,
        headers=headers,
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise AttachmentDownloadError(
            "Vedhæftningen findes ikke længere."
        )

    response.raise_for_status()

    data = response.json()

    attachment_type = _get_attachment_type(
        data
    )

    if attachment_type == "referenceAttachment":
        raise AttachmentDownloadError(
            "referenceAttachment downloades ikke."
        )

    encoded_content = data.get("contentBytes")

    if encoded_content:
        file_bytes = base64.b64decode(
            encoded_content
        )
    else:
        downloaded = _download_attachment_value(
            user_mail=user_mail,
            message_id=message_id,
            attachment_id=attachment_id,
        )

        file_bytes = downloaded[
            "content_bytes"
        ]

    if as_bytes:
        return file_bytes

    if save_path:
        target_path = Path(save_path)

        target_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        target_path.write_bytes(
            file_bytes
        )

        return str(target_path)

    return file_bytes

# -------------------------------------------------
# GET FOLDERS
# -------------------------------------------------

def get_folders(
    user_mail,
    include_hidden=True,
    include_children=True,
):
    """
    Henter postkassens mailmapper.

    include_hidden:
        True medtager skjulte mapper.

    include_children:
        True henter også alle undermapper.

    Returnerer blandt andet:

    {
        "id": "...",
        "display_name": "Autosvar",
        "path": "Indbakke/Autosvar",
        "depth": 1,
        "parent_folder_id": "..."
    }
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    include_hidden_value = (
        "true"
        if include_hidden
        else "false"
    )

    root_url = (
        f"{_user_url(user_mail)}"
        f"/mailFolders"
        f"?includeHiddenFolders="
        f"{include_hidden_value}"
    )

    root_folders = _get_all_pages(
        url=root_url,
        headers=headers,
    )

    formatted_folders = []


    # -------------------------------------------------
    # FORMATÉR ÉN MAPPE
    # -------------------------------------------------

    def format_folder(
        folder,
        path,
        depth,
    ):
        """
        Formaterer én Outlook-mappe.
        """

        display_name = (
            folder.get("displayName")
            or ""
        )

        return {
            # Nye standardfelter.
            "id": folder.get("id"),

            "display_name": (
                display_name
            ),

            "path": path,

            "depth": depth,

            "parent_folder_id": (
                folder.get(
                    "parentFolderId"
                )
            ),

            "child_folder_count": (
                folder.get(
                    "childFolderCount"
                )
                or 0
            ),

            "total_item_count": (
                folder.get(
                    "totalItemCount"
                )
                or 0
            ),

            "unread_item_count": (
                folder.get(
                    "unreadItemCount"
                )
                or 0
            ),

            # Gamle Graph-navne bevares.
            "displayName": (
                display_name
            ),

            "parentFolderId": (
                folder.get(
                    "parentFolderId"
                )
            ),

            "childFolderCount": (
                folder.get(
                    "childFolderCount"
                )
                or 0
            ),

            "totalItemCount": (
                folder.get(
                    "totalItemCount"
                )
                or 0
            ),

            "unreadItemCount": (
                folder.get(
                    "unreadItemCount"
                )
                or 0
            ),
        }


    # -------------------------------------------------
    # HENT UNDERMAPPER
    # -------------------------------------------------

    def get_child_folders(
        parent_folder,
        parent_path,
        depth,
    ):
        """
        Henter alle undermapper rekursivt.

        Rekursiv betyder, at funktionen også
        henter undermappernes egne undermapper.
        """

        parent_folder_id = (
            parent_folder.get("id")
        )

        if not parent_folder_id:
            return

        encoded_folder_id = quote(
            str(parent_folder_id),
            safe="",
        )

        child_url = (
            f"{_user_url(user_mail)}"
            f"/mailFolders/"
            f"{encoded_folder_id}"
            f"/childFolders"
            f"?includeHiddenFolders="
            f"{include_hidden_value}"
        )

        child_folders = _get_all_pages(
            url=child_url,
            headers=headers,
        )

        for child_folder in child_folders:
            child_name = (
                child_folder.get(
                    "displayName"
                )
                or ""
            )

            child_path = (
                f"{parent_path}/"
                f"{child_name}"
            )

            formatted_folders.append(
                format_folder(
                    folder=child_folder,
                    path=child_path,
                    depth=depth,
                )
            )

            child_count = (
                child_folder.get(
                    "childFolderCount"
                )
                or 0
            )

            if child_count > 0:
                get_child_folders(
                    parent_folder=(
                        child_folder
                    ),
                    parent_path=(
                        child_path
                    ),
                    depth=depth + 1,
                )


    # -------------------------------------------------
    # GENNEMGÅ TOPMAPPER
    # -------------------------------------------------

    for root_folder in root_folders:
        root_name = (
            root_folder.get(
                "displayName"
            )
            or ""
        )

        formatted_folders.append(
            format_folder(
                folder=root_folder,
                path=root_name,
                depth=0,
            )
        )

        child_count = (
            root_folder.get(
                "childFolderCount"
            )
            or 0
        )

        if (
            include_children
            and child_count > 0
        ):
            get_child_folders(
                parent_folder=(
                    root_folder
                ),
                parent_path=root_name,
                depth=1,
            )

    return formatted_folders

# -------------------------------------------------
# MARK AS READ
# -------------------------------------------------

def mark_as_read(
    user_mail,
    message_id,
    is_read=True,
):
    """
    Markerer en mail som læst eller ulæst.

    is_read=True:
        Mailen markeres som læst.

    is_read=False:
        Mailen markeres som ulæst.
    """

    headers = _get_headers(
        prefer_plain_text=False,
    )

    response = requests.patch(
        _message_url(user_mail, message_id),
        headers=headers,
        json={
            "isRead": bool(is_read),
        },
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise MailNotFoundError(
            "Mailen findes ikke længere i Outlook."
        )

    response.raise_for_status()

    return True


# -------------------------------------------------
# MOVE MAIL
# -------------------------------------------------

def move_mail(
    user_mail,
    message_id,
    destination_folder_id,
):
    """
    Flytter en mail til en anden mappe.

    destination_folder_id:
        Graph-id eller kendt mappenavn som:
        - archive
        - deleteditems
        - drafts
        - inbox
        - junkemail
        - sentitems
    """

    headers = _get_headers(
        prefer_plain_text=True,
    )

    url = (
        f"{_message_url(user_mail, message_id)}"
        f"/move"
    )

    response = requests.post(
        url,
        headers=headers,
        json={
            "destinationId": (
                destination_folder_id
            ),
        },
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise MailNotFoundError(
            "Mailen eller destinationsmappen "
            "findes ikke."
        )

    response.raise_for_status()

    moved_mail = response.json()

    return _format_mail(
        m=moved_mail,
        user_mail=user_mail,
        folder=destination_folder_id,
        include_attachments=False,
        get_inline=False,
    )


# -------------------------------------------------
# UPDATE MAIL CATEGORIES
# -------------------------------------------------

def update_mail_categories(
    user_mail,
    message_id,
    categories,
):
    """
    Overskriver mailens kategorier.

    categories skal være en liste.

    Eksempel:
    ["Test", "Point: 300"]
    """

    if categories is None:
        categories = []

    if not isinstance(categories, list):
        raise TypeError(
            "categories skal være en liste."
        )

    headers = _get_headers(
        prefer_plain_text=False,
    )

    response = requests.patch(
        _message_url(user_mail, message_id),
        headers=headers,
        json={
            "categories": categories,
        },
        timeout=DEFAULT_TIMEOUT,
    )

    if response.status_code == 404:
        raise MailNotFoundError(
            "Mailen findes ikke længere i Outlook."
        )

    response.raise_for_status()

    return categories


# -------------------------------------------------
# ADD MAIL CATEGORY
# -------------------------------------------------

def add_mail_category(
    user_mail,
    message_id,
    category,
):
    """
    Tilføjer én kategori uden at fjerne
    eksisterende kategorier.
    """

    if not category:
        raise ValueError(
            "category må ikke være tom."
        )

    mails = get_mails(
        user_mail=user_mail,
        message_id=message_id,
        include_attachments=False,
        prefer_plain_text=True,
    )

    if not mails:
        raise MailNotFoundError(
            "Mailen findes ikke længere i Outlook."
        )

    categories = list(
        mails[0].get("categories") or []
    )

    if category not in categories:
        categories.append(category)

    return update_mail_categories(
        user_mail=user_mail,
        message_id=message_id,
        categories=categories,
    )


# -------------------------------------------------
# REMOVE MAIL CATEGORY
# -------------------------------------------------

def remove_mail_category(
    user_mail,
    message_id,
    category,
):
    """
    Fjerner én kategori uden at ændre
    de øvrige kategorier.
    """

    mails = get_mails(
        user_mail=user_mail,
        message_id=message_id,
        include_attachments=False,
        prefer_plain_text=True,
    )

    if not mails:
        raise MailNotFoundError(
            "Mailen findes ikke længere i Outlook."
        )

    categories = [
        existing_category
        for existing_category
        in (
            mails[0].get("categories")
            or []
        )
        if existing_category != category
    ]

    return update_mail_categories(
        user_mail=user_mail,
        message_id=message_id,
        categories=categories,
    )
