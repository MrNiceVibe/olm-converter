from email.message import EmailMessage
from email.policy import SMTP
from typing import Any, Dict, List


def build_email_message(msg: Dict[str, Any]) -> EmailMessage:
    """
    Convert a normalized dict into an EmailMessage.
    """
    em = EmailMessage(policy=SMTP)

    if msg.get("message_id"):
        em["Message-ID"] = msg["message_id"]
    if msg.get("date"):
        em["Date"] = msg["date"]
    em["Subject"] = msg.get("subject") or "(no subject)"

    from_header = _format_address(msg.get("from_name", ""), msg.get("from_addr", ""))
    if from_header:
        em["From"] = from_header

    _set_address_header(em, "To", msg.get("to", []))
    _set_address_header(em, "Cc", msg.get("cc", []))
    _set_address_header(em, "Bcc", msg.get("bcc", []))
    _set_address_header(em, "Reply-To", msg.get("reply_to", []))

    plain = msg.get("body_plain")
    html = msg.get("body_html")

    if plain and html:
        em.set_content(plain)
        em.add_alternative(html, subtype="html")
    elif html:
        em.set_content(html, subtype="html")
    else:
        em.set_content(plain or "")

    for attachment in msg.get("attachments", []):
        _add_attachment(em, attachment)

    return em


def _set_address_header(em: EmailMessage, header: str, values: List[Dict[str, str]]) -> None:
    entries = [_format_address(item.get("name", ""), item.get("addr", "")) for item in values if item.get("addr")]
    if entries:
        em[header] = ", ".join(entries)


def _format_address(name: str, addr: str) -> str:
    addr = (addr or "").strip()
    name = (name or "").strip()
    if not addr:
        return ""
    if name:
        safe_name = name.replace('"', '\\"')
        return f'"{safe_name}" <{addr}>'
    return addr


def _add_attachment(em: EmailMessage, attachment: Dict[str, Any]) -> None:
    content_type = attachment.get("content_type") or "application/octet-stream"
    maintype, _, subtype = content_type.partition("/")
    if not maintype or not subtype:
        maintype, subtype = "application", "octet-stream"

    em.add_attachment(
        attachment.get("data", b""),
        maintype=maintype,
        subtype=subtype,
        filename=attachment.get("filename") or "attachment",
    )

    part = list(em.iter_attachments())[-1]
    content_id = attachment.get("content_id")
    if content_id:
        normalized = content_id if str(content_id).startswith("<") else f"<{content_id}>"
        part["Content-ID"] = normalized
