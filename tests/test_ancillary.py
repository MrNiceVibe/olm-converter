from pathlib import Path

from converter.ancillary import (
    parse_calendar_file,
    parse_contacts_file,
    parse_tasks_file,
    render_calendar_ics,
    render_tasks_ics,
    render_vcf,
)


def test_contacts_export_to_vcf(tmp_path: Path) -> None:
    xml = """<contacts elementCount="1"><contact>
    <OPFContactCopyDisplayName>Marcus Olsson</OPFContactCopyDisplayName>
    <OPFContactCopyFirstName>Marcus</OPFContactCopyFirstName>
    <OPFContactCopyLastName>Olsson</OPFContactCopyLastName>
    <OPFContactCopyBusinessCompany>NiceVibe</OPFContactCopyBusinessCompany>
    <OPFContactCopyBusinessTitle>Founder</OPFContactCopyBusinessTitle>
    <OPFContactCopyCellPhone>+46700000000</OPFContactCopyCellPhone>
    <OPFContactCopyBusinessPhone>+46800000000</OPFContactCopyBusinessPhone>
    <OPFContactCopyBusinessStreetAddress>Main Street 1</OPFContactCopyBusinessStreetAddress>
    <OPFContactCopyBusinessCity>Stockholm</OPFContactCopyBusinessCity>
    <OPFContactCopyBusinessZip>11122</OPFContactCopyBusinessZip>
    <OPFContactCopyBusinessCountry>Sweden</OPFContactCopyBusinessCountry>
    <OPFContactCopyNotesPlain>Hello note</OPFContactCopyNotesPlain>
    <OPFContactCopyEmailAddressList>
      <contactEmailAddress OPFContactEmailAddressAddress="marcus@example.com" OPFContactEmailAddressType="2" />
    </OPFContactCopyEmailAddressList>
    </contact></contacts>"""
    path = tmp_path / "Contacts.xml"
    path.write_text(xml, encoding="utf-8")

    contacts = parse_contacts_file(str(path))
    vcf = render_vcf(contacts)

    assert len(contacts) == 1
    assert "BEGIN:VCARD" in vcf
    assert "FN:Marcus Olsson" in vcf
    assert "EMAIL;TYPE=INTERNET:marcus@example.com" in vcf


def test_calendar_export_to_ics(tmp_path: Path) -> None:
    xml = """<appointments elementCount="1"><appointment>
    <ExchangeServerLastModifiedTime>2025-07-11T14:09:04</ExchangeServerLastModifiedTime>
    <OPFCalendarEventCopyModDate>2025-07-11T14:09:04</OPFCalendarEventCopyModDate>
    <OPFCalendarEventCopyStartTime>2025-07-11T12:15:00</OPFCalendarEventCopyStartTime>
    <OPFCalendarEventCopyEndTime>2025-07-11T14:45:00</OPFCalendarEventCopyEndTime>
    <OPFCalendarEventCopyStartTimeZone>(UTC+07:00) Bangkok, Hanoi, Jakarta</OPFCalendarEventCopyStartTimeZone>
    <OPFCalendarEventCopyEndTimeZone>(UTC+07:00) Bangkok, Hanoi, Jakarta</OPFCalendarEventCopyEndTimeZone>
    <OPFCalendarEventCopySummary>Cool Vibes</OPFCalendarEventCopySummary>
    <OPFCalendarEventCopyDescriptionPlain>Join here</OPFCalendarEventCopyDescriptionPlain>
    <OPFCalendarEventCopyLocation>Online</OPFCalendarEventCopyLocation>
    <OPFCalendarEventCopyOrganizer>marcus@example.com</OPFCalendarEventCopyOrganizer>
    <OPFCalendarEventCopyUUID>event-123</OPFCalendarEventCopyUUID>
    <OPFCalendarEventGetIsAllDayEvent>0</OPFCalendarEventGetIsAllDayEvent>
    <OPFCalendarEventGetHasReminder>1</OPFCalendarEventGetHasReminder>
    <OPFCalendarEventCopyReminderDelta>1800</OPFCalendarEventCopyReminderDelta>
    <OPFCalendarEventGetStartTimeZoneICSData>BEGIN:VCALENDAR
BEGIN:VTIMEZONE
TZID:SE Asia Standard Time
END:VTIMEZONE
END:VCALENDAR</OPFCalendarEventGetStartTimeZoneICSData>
    </appointment></appointments>"""
    path = tmp_path / "Calendar.xml"
    path.write_text(xml, encoding="utf-8")

    calendar = parse_calendar_file(str(path))
    ics = render_calendar_ics(calendar, "Calendar")

    assert len(calendar["events"]) == 1
    assert "BEGIN:VEVENT" in ics
    assert "SUMMARY:Cool Vibes" in ics
    assert "LOCATION:Online" in ics
    assert "BEGIN:VALARM" in ics


def test_tasks_export_to_ics(tmp_path: Path) -> None:
    xml = """<tasks elementCount="1"><task>
    <OPFTaskCopyModDate>2025-07-24T09:43:46</OPFTaskCopyModDate>
    <OPFTaskCopyCompletedDateTime>2025-07-14T00:00:00</OPFTaskCopyCompletedDateTime>
    <OPFTaskCopyName>Prep for launch</OPFTaskCopyName>
    <OPFTaskCopyNotePlain>Remember the checklist</OPFTaskCopyNotePlain>
    <OPFTaskCopyUUID>task-123</OPFTaskCopyUUID>
    </task></tasks>"""
    path = tmp_path / "Tasks.xml"
    path.write_text(xml, encoding="utf-8")

    tasks = parse_tasks_file(str(path))
    ics = render_tasks_ics(tasks, "Tasks")

    assert len(tasks["todos"]) == 1
    assert "BEGIN:VTODO" in ics
    assert "SUMMARY:Prep for launch" in ics
    assert "STATUS:COMPLETED" in ics
