import html
import re
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Dict, List, Optional
import xml.etree.ElementTree as ET

import chardet


EMAIL_ADDRESS_ATTR = "OPFContactEmailAddressAddress"
EMAIL_NAME_ATTR = "OPFContactEmailAddressName"


class _HTMLToTextParser(HTMLParser):
    BLOCK_TAGS = {
        "address",
        "article",
        "aside",
        "blockquote",
        "br",
        "div",
        "dl",
        "dt",
        "dd",
        "fieldset",
        "figcaption",
        "figure",
        "footer",
        "form",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "header",
        "hr",
        "li",
        "main",
        "ol",
        "p",
        "pre",
        "section",
        "table",
        "tr",
        "td",
        "th",
        "ul",
    }

    SKIP_TAGS = {"script", "style", "head", "title"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._parts: List[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in self.SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        if tag in self.BLOCK_TAGS:
            self._parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in self.SKIP_TAGS and self._skip_depth:
            self._skip_depth -= 1
            return
        if self._skip_depth:
            return
        if tag in self.BLOCK_TAGS:
            self._parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip_depth and data:
            self._parts.append(data)

    def get_text(self) -> str:
        text = "".join(self._parts)
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = re.sub(r"\n[ \t]+", "\n", text)
        text = re.sub(r"[ \t]+\n", "\n", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()


def parse_message_file(filepath: str) -> Dict[str, Any]:
    """
    Parse a real Outlook for Mac message XML file into a normalized structure.
    """
    raw = _read_file(filepath)
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise ValueError(f"XML parse error in {filepath}: {exc}") from exc

    email_el = _get_email_element(root)

    from_values = _get_addresses(email_el, "OPFMessageCopyFromAddresses")
    sender_values = _get_addresses(email_el, "OPFMessageCopySenderAddress")
    from_entry = (from_values or sender_values or [{"name": "", "addr": ""}])[0]

    body_plain, body_html = _extract_bodies(email_el)
    attachments = _get_attachments(email_el, filepath)

    return {
        "message_id": _normalize_message_id(_get_text(email_el, "OPFMessageCopyMessageID")),
        "from_name": from_entry["name"],
        "from_addr": from_entry["addr"],
        "to": _get_addresses(email_el, "OPFMessageCopyToAddresses"),
        "cc": _get_addresses(email_el, "OPFMessageCopyCCAddresses"),
        "bcc": _get_addresses(email_el, "OPFMessageCopyBCCAddresses"),
        "reply_to": _get_addresses(email_el, "OPFMessageCopyReplyToAddresses"),
        "subject": _get_text(email_el, "OPFMessageCopySubject") or "(no subject)",
        "date": _get_date(email_el),
        "body_plain": body_plain,
        "body_html": body_html,
        "attachments": attachments,
        "_source_file": filepath,
    }


def _read_file(filepath: str) -> str:
    raw_bytes = Path(filepath).read_bytes()
    detected = chardet.detect(raw_bytes)
    encoding = detected.get("encoding") or "utf-8"
    try:
        return raw_bytes.decode(encoding)
    except (UnicodeDecodeError, LookupError):
        return raw_bytes.decode("utf-8", errors="replace")


def _get_email_element(root: ET.Element) -> ET.Element:
    if root.tag == "email":
        return root
    email_el = root.find("./email")
    if email_el is not None:
        return email_el
    raise ValueError("Message XML does not contain an <email> element.")


def _get_text(parent: ET.Element, tag: str) -> Optional[str]:
    el = parent.find(f"./{tag}")
    if el is None:
        return None
    text = "".join(el.itertext()).strip()
    return text or None


def _get_addresses(parent: ET.Element, tag: str) -> List[Dict[str, str]]:
    container = parent.find(f"./{tag}")
    if container is None:
        return []

    results: List[Dict[str, str]] = []
    for item in list(container.findall("./emailAddress")) + list(container.findall("./contactEmailAddress")):
        addr = (item.get(EMAIL_ADDRESS_ATTR) or (item.text or "")).strip()
        name = (item.get(EMAIL_NAME_ATTR) or "").strip()
        if addr:
            results.append({"name": name, "addr": addr})

    if not results:
        direct = "".join(container.itertext()).strip()
        if direct:
            results.append({"name": "", "addr": direct})

    return results


def _extract_bodies(parent: ET.Element) -> tuple[Optional[str], Optional[str]]:
    raw_body = _get_text(parent, "OPFMessageCopyBody")
    raw_html = _get_text(parent, "OPFMessageCopyHTMLBody")
    preview = _get_text(parent, "OPFMessageCopyPreview")

    body_plain: Optional[str] = None
    body_html: Optional[str] = None

    if raw_html:
        body_html = raw_html
        if raw_body and not _looks_like_html(raw_body):
            body_plain = _cleanup_text_body(raw_body)
        else:
            body_plain = _html_to_text(raw_html)
    elif raw_body:
        if _looks_like_html(raw_body):
            body_html = raw_body
            body_plain = _html_to_text(raw_body)
        else:
            body_plain = _cleanup_text_body(raw_body)

    if not body_plain and preview:
        body_plain = _cleanup_text_body(preview)

    return body_plain or None, body_html or None


def _cleanup_text_body(text: str) -> str:
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n")
    cleaned = cleaned.replace("\xa0", " ")
    cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def _looks_like_html(text: str) -> bool:
    return bool(re.search(r"<(?:!DOCTYPE|html|body|div|p|br|table|span|head|meta)\b", text, re.IGNORECASE))


def _html_to_text(value: str) -> str:
    parser = _HTMLToTextParser()
    parser.feed(value)
    parser.close()
    return parser.get_text()


def _get_date(parent: ET.Element) -> str:
    raw = (
        _get_text(parent, "OPFMessageCopyReceivedTime")
        or _get_text(parent, "OPFMessageCopySentTime")
        or _get_text(parent, "ExchangeServerLastModifiedTime")
    )
    if not raw:
        return "Thu, 01 Jan 1970 00:00:00 -0000"

    for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S"):
        try:
            dt = datetime.strptime(raw, fmt)
            if fmt.endswith("%z"):
                return dt.strftime("%a, %d %b %Y %H:%M:%S %z")
            if raw.endswith("Z"):
                return dt.strftime("%a, %d %b %Y %H:%M:%S +0000")
            return dt.strftime("%a, %d %b %Y %H:%M:%S -0000")
        except ValueError:
            continue

    return "Thu, 01 Jan 1970 00:00:00 -0000"


def _get_attachments(parent: ET.Element, filepath: str) -> List[Dict[str, Any]]:
    attachments: List[Dict[str, Any]] = []
    seen_sources: set[str] = set()

    container = parent.find("./OPFMessageCopyAttachmentList")
    if container is not None:
        for item in container.findall("./messageAttachment"):
            rel_url = html.unescape(item.get("OPFAttachmentURL", "")).strip()
            source_path = _resolve_archive_path(filepath, rel_url) if rel_url else None
            if source_path is not None:
                seen_sources.add(str(source_path))
            attachments.append(
                {
                    "filename": item.get("OPFAttachmentName") or (source_path.name if source_path else "attachment"),
                    "content_type": item.get("OPFAttachmentContentType") or "application/octet-stream",
                    "content_id": item.get("OPFAttachmentContentID"),
                    "disposition": "attachment",
                    "data": source_path.read_bytes() if source_path and source_path.exists() else b"",
                    "source_path": str(source_path) if source_path else None,
                }
            )

    meeting_data = _get_text(parent, "OPFMessageCopyMeetingData")
    if meeting_data:
        source_path = _resolve_archive_path(filepath, html.unescape(meeting_data))
        if source_path is not None and str(source_path) not in seen_sources:
            attachments.append(
                {
                    "filename": source_path.name,
                    "content_type": "text/calendar",
                    "content_id": None,
                    "disposition": "attachment",
                    "data": source_path.read_bytes() if source_path.exists() else b"",
                    "source_path": str(source_path),
                }
            )

    return attachments


def _resolve_archive_path(message_path: str, rel_url: str) -> Optional[Path]:
    message_file = Path(message_path).resolve()
    accounts_root = _find_accounts_root(message_file)
    if accounts_root is None:
        return None
    archive_root = accounts_root.parent
    candidate = (archive_root / Path(rel_url)).resolve()
    return candidate


def _find_accounts_root(path: Path) -> Optional[Path]:
    for parent in (path.parent,) + tuple(path.parents):
        if parent.name == "Accounts":
            return parent
    return None


def _normalize_message_id(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    value = value.strip()
    if value.startswith("<") and value.endswith(">"):
        return value
    if "@" in value:
        return f"<{value}>"
    return value
