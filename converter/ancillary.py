import base64
import hashlib
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional
import xml.etree.ElementTree as ET

import chardet


def parse_contacts_file(filepath: str) -> List[Dict[str, Any]]:
    root = _read_xml_root(filepath)
    contacts: List[Dict[str, Any]] = []

    for item in root.findall("./contact"):
        first_name = _get_text(item, "OPFContactCopyFirstName") or ""
        last_name = _get_text(item, "OPFContactCopyLastName") or ""
        display_name = _get_text(item, "OPFContactCopyDisplayName") or "Unnamed Contact"
        notes = _coalesce_text(
            _get_text(item, "OPFContactCopyNotesPlain"),
            _html_to_text(_get_text(item, "OPFContactCopyNotes")),
        )

        emails = _get_contact_emails(item)
        contact = {
            "uid": _get_text(item, "OPFContactExchangeID") or _stable_uid(filepath, display_name),
            "display_name": display_name,
            "first_name": first_name,
            "last_name": last_name,
            "emails": emails,
            "organization": _get_text(item, "OPFContactCopyBusinessCompany"),
            "title": _get_text(item, "OPFContactCopyBusinessTitle"),
            "cell_phone": _get_text(item, "OPFContactCopyCellPhone"),
            "business_phone": _get_text(item, "OPFContactCopyBusinessPhone"),
            "notes": notes,
            "address": {
                "street": _get_text(item, "OPFContactCopyBusinessStreetAddress"),
                "city": _get_text(item, "OPFContactCopyBusinessCity"),
                "state": _get_text(item, "OPFContactCopyBusinessState"),
                "postal_code": _get_text(item, "OPFContactCopyBusinessZip"),
                "country": _get_text(item, "OPFContactCopyBusinessCountry"),
            },
            "photo_b64": _get_text(item, "OPFContactAddPicture"),
        }
        contacts.append(contact)

    return contacts


def render_vcf(contacts: List[Dict[str, Any]]) -> str:
    lines: List[str] = []
    for contact in contacts:
        lines.extend(
            [
                "BEGIN:VCARD",
                "VERSION:3.0",
                _fold_vcard_line(f"UID:{_escape_text(contact['uid'])}"),
                _fold_vcard_line(
                    f"N:{_escape_text(contact['last_name'])};{_escape_text(contact['first_name'])};;;"
                ),
                _fold_vcard_line(f"FN:{_escape_text(contact['display_name'])}"),
            ]
        )

        if contact.get("organization"):
            lines.append(_fold_vcard_line(f"ORG:{_escape_text(contact['organization'])}"))
        if contact.get("title"):
            lines.append(_fold_vcard_line(f"TITLE:{_escape_text(contact['title'])}"))
        if contact.get("business_phone"):
            lines.append(_fold_vcard_line(f"TEL;TYPE=WORK:{_escape_text(contact['business_phone'])}"))
        if contact.get("cell_phone"):
            lines.append(_fold_vcard_line(f"TEL;TYPE=CELL:{_escape_text(contact['cell_phone'])}"))
        if contact.get("notes"):
            lines.append(_fold_vcard_line(f"NOTE:{_escape_text(contact['notes'])}"))

        address = contact.get("address", {})
        if any(address.values()):
            lines.append(
                _fold_vcard_line(
                    "ADR;TYPE=WORK:;;{street};{city};{state};{postal_code};{country}".format(
                        street=_escape_text(address.get("street") or ""),
                        city=_escape_text(address.get("city") or ""),
                        state=_escape_text(address.get("state") or ""),
                        postal_code=_escape_text(address.get("postal_code") or ""),
                        country=_escape_text(address.get("country") or ""),
                    )
                )
            )

        for email in contact.get("emails", []):
            email_type = email.get("type") or "INTERNET"
            lines.append(
                _fold_vcard_line(
                    f"EMAIL;TYPE={email_type}:{_escape_text(email.get('address') or '')}"
                )
            )

        if contact.get("photo_b64"):
            lines.append(_fold_vcard_line(f"PHOTO;ENCODING=b;TYPE=PNG:{contact['photo_b64']}"))

        lines.append("END:VCARD")

    return "\r\n".join(lines) + ("\r\n" if lines else "")


def parse_calendar_file(filepath: str) -> Dict[str, Any]:
    root = _read_xml_root(filepath)
    events: List[Dict[str, Any]] = []
    timezones: Dict[str, str] = {}

    for item in root.findall("./appointment"):
        tz_block = _get_text(item, "OPFCalendarEventGetStartTimeZoneICSData")
        tzid = _extract_tzid(tz_block)
        if tzid and tz_block:
            timezones[tzid] = tz_block

        start = _parse_calendar_datetime(
            _get_text(item, "OPFCalendarEventCopyStartTime"),
            _get_text(item, "OPFCalendarEventCopyStartTimeZone"),
        )
        end = _parse_calendar_datetime(
            _get_text(item, "OPFCalendarEventCopyEndTime"),
            _get_text(item, "OPFCalendarEventCopyEndTimeZone"),
        )
        is_all_day = (_get_text(item, "OPFCalendarEventGetIsAllDayEvent") or "0").startswith("1")

        events.append(
            {
                "uid": _get_text(item, "OPFCalendarEventCopyUUID") or _stable_uid(filepath, str(start)),
                "summary": _get_text(item, "OPFCalendarEventCopySummary") or "(no title)",
                "description": _coalesce_text(
                    _get_text(item, "OPFCalendarEventCopyDescriptionPlain"),
                    _html_to_text(_get_text(item, "OPFCalendarEventCopyDescription")),
                ),
                "location": _get_text(item, "OPFCalendarEventCopyLocation"),
                "organizer": _get_text(item, "OPFCalendarEventCopyOrganizer"),
                "start": start,
                "end": end,
                "is_all_day": is_all_day,
                "tzid": tzid,
                "rrule": _build_rrule(item.find("./OPFCalendarEventCopyRecurrence")),
                "reminder_seconds": _parse_int(_get_text(item, "OPFCalendarEventCopyReminderDelta")),
                "has_reminder": (_get_text(item, "OPFCalendarEventGetHasReminder") or "0").startswith("1"),
                "created": _parse_basic_datetime(_get_text(item, "ExchangeServerLastModifiedTime")),
                "modified": _parse_basic_datetime(_get_text(item, "OPFCalendarEventCopyModDate")),
            }
        )

    return {"events": events, "timezones": timezones}


def render_calendar_ics(calendar: Dict[str, Any], calendar_name: str) -> str:
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//OLM Converter//EN",
        "CALSCALE:GREGORIAN",
        _fold_ical_line(f"X-WR-CALNAME:{_escape_ical(calendar_name)}"),
    ]

    for tz_block in calendar.get("timezones", {}).values():
        lines.extend(_normalize_ical_block(tz_block))

    for event in calendar.get("events", []):
        lines.append("BEGIN:VEVENT")
        lines.append(_fold_ical_line(f"UID:{_escape_ical(event['uid'])}"))
        lines.append(_fold_ical_line(f"SUMMARY:{_escape_ical(event['summary'])}"))
        if event.get("description"):
            lines.append(_fold_ical_line(f"DESCRIPTION:{_escape_ical(event['description'])}"))
        if event.get("location"):
            lines.append(_fold_ical_line(f"LOCATION:{_escape_ical(event['location'])}"))
        if event.get("organizer"):
            lines.append(_fold_ical_line(f"ORGANIZER:mailto:{event['organizer']}"))
        if event.get("created"):
            lines.append(_fold_ical_line(f"CREATED:{event['created'].strftime('%Y%m%dT%H%M%SZ')}"))
        if event.get("modified"):
            lines.append(_fold_ical_line(f"LAST-MODIFIED:{event['modified'].strftime('%Y%m%dT%H%M%SZ')}"))

        _append_dt_lines(lines, "DTSTART", event["start"], event.get("tzid"), event["is_all_day"])
        if event.get("end"):
            _append_dt_lines(lines, "DTEND", event["end"], event.get("tzid"), event["is_all_day"])
        if event.get("rrule"):
            lines.append(_fold_ical_line(f"RRULE:{event['rrule']}"))
        if event.get("has_reminder") and event.get("reminder_seconds"):
            lines.extend(
                [
                    "BEGIN:VALARM",
                    _fold_ical_line(f"TRIGGER:-PT{int(event['reminder_seconds'])}S"),
                    "ACTION:DISPLAY",
                    _fold_ical_line(f"DESCRIPTION:{_escape_ical(event['summary'])}"),
                    "END:VALARM",
                ]
            )
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def parse_tasks_file(filepath: str) -> Dict[str, Any]:
    root = _read_xml_root(filepath)
    todos: List[Dict[str, Any]] = []

    for item in root.findall("./task"):
        completed = _parse_basic_datetime(_get_text(item, "OPFTaskCopyCompletedDateTime"))
        due = _parse_basic_datetime(_coalesce_text(_get_text(item, "OPFTaskCopyDueDate"), _get_text(item, "OPFTaskCopyDueDateTime")))
        start = _parse_basic_datetime(_coalesce_text(_get_text(item, "OPFTaskCopyStartDate"), _get_text(item, "OPFTaskCopyStartDateTime")))
        modified = _parse_basic_datetime(_get_text(item, "OPFTaskCopyModDate"))
        summary = _get_text(item, "OPFTaskCopyName") or "(no task title)"
        description = _coalesce_text(
            _get_text(item, "OPFTaskCopyNotePlain"),
            _html_to_text(_get_text(item, "OPFTaskCopyNote")),
        )
        todos.append(
            {
                "uid": _get_text(item, "OPFTaskCopyUUID") or _stable_uid(filepath, summary),
                "summary": summary,
                "description": description,
                "start": start,
                "due": due,
                "completed": completed,
                "modified": modified,
                "status": "COMPLETED" if completed else "NEEDS-ACTION",
            }
        )

    return {"todos": todos}


def render_tasks_ics(tasks: Dict[str, Any], list_name: str) -> str:
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//OLM Converter//EN",
        "CALSCALE:GREGORIAN",
        _fold_ical_line(f"X-WR-CALNAME:{_escape_ical(list_name)}"),
    ]

    for todo in tasks.get("todos", []):
        lines.append("BEGIN:VTODO")
        lines.append(_fold_ical_line(f"UID:{_escape_ical(todo['uid'])}"))
        lines.append(_fold_ical_line(f"SUMMARY:{_escape_ical(todo['summary'])}"))
        lines.append(_fold_ical_line(f"STATUS:{todo['status']}"))
        if todo.get("description"):
            lines.append(_fold_ical_line(f"DESCRIPTION:{_escape_ical(todo['description'])}"))
        if todo.get("modified"):
            lines.append(_fold_ical_line(f"LAST-MODIFIED:{todo['modified'].strftime('%Y%m%dT%H%M%SZ')}"))
        if todo.get("completed"):
            lines.append(_fold_ical_line(f"COMPLETED:{todo['completed'].strftime('%Y%m%dT%H%M%SZ')}"))
        if todo.get("start"):
            _append_dt_lines(lines, "DTSTART", todo["start"], None, False)
        if todo.get("due"):
            _append_dt_lines(lines, "DUE", todo["due"], None, False)
        lines.append("END:VTODO")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def _read_xml_root(filepath: str) -> ET.Element:
    raw_bytes = Path(filepath).read_bytes()
    detected = chardet.detect(raw_bytes)
    encoding = detected.get("encoding") or "utf-8"
    try:
        text = raw_bytes.decode(encoding)
    except (UnicodeDecodeError, LookupError):
        text = raw_bytes.decode("utf-8", errors="replace")
    try:
        return ET.fromstring(text)
    except ET.ParseError as exc:
        raise ValueError(f"XML parse error in {filepath}: {exc}") from exc


def _get_text(parent: ET.Element, tag: str) -> Optional[str]:
    element = parent.find(f"./{tag}")
    if element is None:
        return None
    text = "".join(element.itertext()).strip()
    return text or None


def _get_contact_emails(item: ET.Element) -> List[Dict[str, str]]:
    emails: List[Dict[str, str]] = []
    for container_name in (
        "OPFContactCopyEmailAddressList",
        "OPFContactCopyEmailAddressList1",
        "OPFContactCopyEmailAddressList2",
        "OPFContactCopyDefaultEmailAddress",
    ):
        container = item.find(f"./{container_name}")
        if container is None:
            continue
        for email in container.findall("./contactEmailAddress"):
            address = (email.get("OPFContactEmailAddressAddress") or "").strip()
            if not address:
                continue
            email_type = email.get("OPFContactEmailAddressType") or "INTERNET"
            if address not in {entry["address"] for entry in emails}:
                emails.append({"address": address, "type": _normalize_email_type(email_type)})
    return emails


def _normalize_email_type(value: str) -> str:
    mapping = {"0": "WORK", "1": "HOME", "2": "INTERNET", "3": "X-IM"}
    return mapping.get(value, value or "INTERNET")


def _html_to_text(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    text = re.sub(r"(?i)<br\s*/?>", "\n", value)
    text = re.sub(r"(?i)</p>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&nbsp;", " ")
    text = re.sub(r"\r\n?", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() or None


def _coalesce_text(*values: Optional[str]) -> Optional[str]:
    for value in values:
        if value:
            return value.strip()
    return None


def _stable_uid(filepath: str, seed: str) -> str:
    digest = hashlib.sha1(f"{filepath}:{seed}".encode("utf-8")).hexdigest()
    return f"olm-converter-{digest}"


def _parse_basic_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def _parse_calendar_datetime(value: Optional[str], timezone_label: Optional[str]) -> Optional[datetime]:
    dt = _parse_basic_datetime(value)
    if dt is None:
        return None
    offset = _parse_timezone_offset(timezone_label)
    if offset is not None:
        return dt + offset
    return dt


def _parse_timezone_offset(label: Optional[str]) -> Optional[timedelta]:
    if not label:
        return None
    match = re.search(r"\(UTC([+-])(\d{2}):(\d{2})\)", label)
    if not match:
        return None
    sign = 1 if match.group(1) == "+" else -1
    hours = int(match.group(2))
    minutes = int(match.group(3))
    return timedelta(hours=sign * hours, minutes=sign * minutes)


def _extract_tzid(tz_block: Optional[str]) -> Optional[str]:
    if not tz_block:
        return None
    match = re.search(r"^TZID:(.+)$", tz_block, re.MULTILINE)
    return match.group(1).strip() if match else None


def _build_rrule(recurrence: Optional[ET.Element]) -> Optional[str]:
    if recurrence is None:
        return None
    pattern = recurrence.find("./OPFRecurrencePattern")
    if pattern is None:
        return None

    pattern_type = _get_text(pattern, "OPFRecurrencePatternType")
    interval = _get_text(pattern, "OPFRecurrencePatternInterval") or "1"
    if not pattern_type:
        return None

    parts = []
    mapping = {
        "OPFRecurrencePatternDaily": "FREQ=DAILY",
        "OPFRecurrencePatternWeekly": "FREQ=WEEKLY",
        "OPFRecurrencePatternAbsoluteMonthly": "FREQ=MONTHLY",
        "OPFRecurrencePatternAbsoluteYearly": "FREQ=YEARLY",
    }
    if pattern_type not in mapping:
        return None
    parts.append(mapping[pattern_type])
    parts.append(f"INTERVAL={interval}")

    if pattern_type == "OPFRecurrencePatternWeekly":
        days = []
        day_mapping = {
            "Sunday": "SU",
            "Monday": "MO",
            "Tuesday": "TU",
            "Wednesday": "WE",
            "Thursday": "TH",
            "Friday": "FR",
            "Saturday": "SA",
        }
        for xml_name, code in day_mapping.items():
            value = _get_text(pattern, f"OPFRecurrencePattern{xml_name}")
            if value and value.startswith("1"):
                days.append(code)
        if days:
            parts.append(f"BYDAY={','.join(days)}")

    if pattern_type in {"OPFRecurrencePatternAbsoluteMonthly", "OPFRecurrencePatternAbsoluteYearly"}:
        day_of_month = _get_text(pattern, "OPFRecurrencePatternDayOfMonth")
        if day_of_month:
            parts.append(f"BYMONTHDAY={day_of_month}")

    if pattern_type == "OPFRecurrencePatternAbsoluteYearly":
        month = _get_text(pattern, "OPFRecurrencePatternMonth")
        if month:
            parts.append(f"BYMONTH={month}")

    return ";".join(parts)


def _append_dt_lines(lines: List[str], field_name: str, value: datetime, tzid: Optional[str], is_all_day: bool) -> None:
    if is_all_day:
        lines.append(_fold_ical_line(f"{field_name};VALUE=DATE:{value.strftime('%Y%m%d')}"))
    elif tzid:
        lines.append(_fold_ical_line(f"{field_name};TZID={tzid}:{value.strftime('%Y%m%dT%H%M%S')}"))
    else:
        lines.append(_fold_ical_line(f"{field_name}:{value.strftime('%Y%m%dT%H%M%S')}"))


def _normalize_ical_block(value: str) -> List[str]:
    lines = value.replace("\r\n", "\n").replace("\r", "\n").strip().split("\n")
    return [line.rstrip() for line in lines if line.strip()]


def _parse_int(value: Optional[str]) -> Optional[int]:
    if not value:
        return None
    match = re.search(r"-?\d+", value.replace(",", ""))
    return int(match.group(0)) if match else None


def _escape_text(value: str) -> str:
    value = value.replace("\\", "\\\\")
    value = value.replace("\n", "\\n")
    value = value.replace(";", r"\;")
    value = value.replace(",", r"\,")
    return value


def _escape_ical(value: str) -> str:
    return _escape_text(value)


def _fold_ical_line(value: str, width: int = 75) -> str:
    if len(value) <= width:
        return value
    parts = [value[:width]]
    value = value[width:]
    while value:
        parts.append(" " + value[: width - 1])
        value = value[width - 1 :]
    return "\r\n".join(parts)


def _fold_vcard_line(value: str, width: int = 75) -> str:
    return _fold_ical_line(value, width=width)
