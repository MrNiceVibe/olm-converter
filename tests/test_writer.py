import mailbox
from email import message_from_bytes
from pathlib import Path

from converter.extractor import extract
from converter.writer import write_all


def _write_message(path: Path, sender: str, subject: str, body: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"""<emails><email>
        <OPFMessageCopySenderAddress>
          <emailAddress OPFContactEmailAddressAddress="{sender}" OPFContactEmailAddressName="Sender" />
        </OPFMessageCopySenderAddress>
        <OPFMessageCopySubject>{subject}</OPFMessageCopySubject>
        <OPFMessageCopyReceivedTime>2025-06-01T14:22:00</OPFMessageCopyReceivedTime>
        <OPFMessageCopyBody>{body}</OPFMessageCopyBody>
        </email></emails>""",
        encoding="utf-8",
    )


def _write_file(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def test_write_all_creates_eml_and_mbox_outputs(tmp_path: Path) -> None:
    archive_root = tmp_path / "archive"
    _write_message(
        archive_root / "Accounts" / "first@example.com" / "com.microsoft.__Messages" / "Inbox" / "message_00001.xml",
        "first@example.com",
        "Inbox message",
        "Hello inbox",
    )
    _write_message(
        archive_root / "Accounts" / "second@example.com" / "com.microsoft.__Messages" / "Projects" / "message_00001.xml",
        "second@example.com",
        "Project message",
        "Hello project",
    )

    result = extract(str(archive_root))
    out_dir = tmp_path / "export"
    summary = write_all(result, str(out_dir), {"mbox", "eml"})

    assert summary["written"] == 2
    assert summary["errors"] == []

    eml_files = sorted((out_dir / "eml").rglob("*.eml"))
    assert len(eml_files) == 2
    for eml in eml_files:
        assert message_from_bytes(eml.read_bytes())["Subject"] is not None

    total_messages = 0
    for mbox_path in (out_dir / "mbox").rglob("*.mbox"):
        total_messages += len(mailbox.mbox(str(mbox_path)))
        assert mbox_path.read_bytes().startswith(b"From ")
    assert total_messages == 2


def test_write_all_creates_supporting_contact_calendar_and_task_outputs(tmp_path: Path) -> None:
    archive_root = tmp_path / "archive"
    account_root = archive_root / "Accounts" / "account@example.com"

    _write_message(
        account_root / "com.microsoft.__Messages" / "Inbox" / "message_00001.xml",
        "sender@example.com",
        "Hello world",
        "This is the message body",
    )
    _write_file(
        account_root / "Contacts" / "Contacts.xml",
        """<contacts elementCount="1"><contact>
        <OPFContactCopyDisplayName>Marcus Olsson</OPFContactCopyDisplayName>
        <OPFContactCopyFirstName>Marcus</OPFContactCopyFirstName>
        <OPFContactCopyLastName>Olsson</OPFContactCopyLastName>
        <OPFContactCopyEmailAddressList>
          <contactEmailAddress OPFContactEmailAddressAddress="marcus@example.com" OPFContactEmailAddressType="2" />
        </OPFContactCopyEmailAddressList>
        </contact></contacts>""",
    )
    _write_file(
        account_root / "Calendar" / "Calendar.xml",
        """<appointments elementCount="1"><appointment>
        <OPFCalendarEventCopyStartTime>2025-07-11T12:15:00</OPFCalendarEventCopyStartTime>
        <OPFCalendarEventCopyEndTime>2025-07-11T14:45:00</OPFCalendarEventCopyEndTime>
        <OPFCalendarEventCopySummary>Cool Vibes</OPFCalendarEventCopySummary>
        <OPFCalendarEventCopyDescriptionPlain>Join here</OPFCalendarEventCopyDescriptionPlain>
        <OPFCalendarEventCopyLocation>Online</OPFCalendarEventCopyLocation>
        <OPFCalendarEventCopyOrganizer>marcus@example.com</OPFCalendarEventCopyOrganizer>
        <OPFCalendarEventCopyUUID>event-123</OPFCalendarEventCopyUUID>
        <OPFCalendarEventGetIsAllDayEvent>0</OPFCalendarEventGetIsAllDayEvent>
        </appointment></appointments>""",
    )
    _write_file(
        account_root / "Tasks" / "Tasks.xml",
        """<tasks elementCount="1"><task>
        <OPFTaskCopyModDate>2025-07-24T09:43:46</OPFTaskCopyModDate>
        <OPFTaskCopyName>Prep for launch</OPFTaskCopyName>
        <OPFTaskCopyNotePlain>Remember the checklist</OPFTaskCopyNotePlain>
        <OPFTaskCopyUUID>task-123</OPFTaskCopyUUID>
        </task></tasks>""",
    )

    result = extract(str(archive_root))
    out_dir = tmp_path / "export"
    summary = write_all(result, str(out_dir), {"mbox"})

    assert summary["written"] == 1
    assert summary["contacts_written"] == 1
    assert summary["calendar_files_written"] == 1
    assert summary["task_files_written"] == 1
    assert summary["errors"] == []

    mbox_path = out_dir / "mbox" / "account@example.com" / "Inbox.mbox"
    contacts_path = out_dir / "contacts" / "account@example.com" / "Contacts.vcf"
    calendar_path = out_dir / "calendars" / "account@example.com" / "Calendar.ics"
    tasks_path = out_dir / "tasks" / "account@example.com" / "Tasks.ics"

    assert mbox_path.exists()
    assert contacts_path.exists()
    assert calendar_path.exists()
    assert tasks_path.exists()

    mbox_messages = mailbox.mbox(str(mbox_path))
    assert len(mbox_messages) == 1
    assert mbox_messages[0]["Subject"] == "Hello world"

    contacts_export = contacts_path.read_text(encoding="utf-8")
    calendar_export = calendar_path.read_text(encoding="utf-8")
    tasks_export = tasks_path.read_text(encoding="utf-8")

    assert "BEGIN:VCARD" in contacts_export
    assert "FN:Marcus Olsson" in contacts_export
    assert "marcus@example.com" in contacts_export

    assert "BEGIN:VEVENT" in calendar_export
    assert "SUMMARY:Cool Vibes" in calendar_export
    assert "LOCATION:Online" in calendar_export

    assert "BEGIN:VTODO" in tasks_export
    assert "SUMMARY:Prep for launch" in tasks_export
    assert "Remember the checklist" in tasks_export
