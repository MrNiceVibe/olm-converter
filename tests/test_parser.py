from pathlib import Path

from converter.parser import parse_message_file


SAMPLE_XML = """<emails xml:space="preserve" elementCount="1">
  <email xml:space="preserve">
    <OPFMessageCopyMessageID>abc123@example.com</OPFMessageCopyMessageID>
    <OPFMessageCopyFromAddresses>
      <emailAddress OPFContactEmailAddressAddress="sender@example.com" OPFContactEmailAddressName="Sender Name" />
    </OPFMessageCopyFromAddresses>
    <OPFMessageCopyToAddresses>
      <emailAddress OPFContactEmailAddressAddress="alice@example.com" OPFContactEmailAddressName="Alice" />
      <emailAddress OPFContactEmailAddressAddress="bob@example.com" OPFContactEmailAddressName="Bob" />
    </OPFMessageCopyToAddresses>
    <OPFMessageCopyCCAddresses>
      <emailAddress OPFContactEmailAddressAddress="carol@example.com" OPFContactEmailAddressName="Carol" />
    </OPFMessageCopyCCAddresses>
    <OPFMessageCopySubject>Hello world</OPFMessageCopySubject>
    <OPFMessageCopyReceivedTime>2025-06-01T14:22:00</OPFMessageCopyReceivedTime>
    <OPFMessageCopyBody>&lt;html&gt;&lt;body&gt;&lt;p&gt;Hello plain&lt;/p&gt;&lt;/body&gt;&lt;/html&gt;</OPFMessageCopyBody>
    <OPFMessageCopyHTMLBody>&lt;html&gt;&lt;body&gt;&lt;p&gt;Hello plain&lt;/p&gt;&lt;/body&gt;&lt;/html&gt;</OPFMessageCopyHTMLBody>
    <OPFMessageCopyAttachmentList>
      <messageAttachment
        OPFAttachmentName="invoice.pdf"
        OPFAttachmentContentType="application/pdf"
        OPFAttachmentContentID="cid-123"
        OPFAttachmentURL="Accounts/account@example.com/com.microsoft.__Messages/Inbox/com.microsoft.__Attachments/invoice_0000" />
    </OPFMessageCopyAttachmentList>
    <OPFMessageCopyMeetingData>Accounts/account@example.com/com.microsoft.__Messages/Inbox/com.microsoft.__Attachments/meeting.ics</OPFMessageCopyMeetingData>
  </email>
</emails>
"""


def test_parse_real_outlook_message_shape(tmp_path: Path) -> None:
    root = tmp_path / "archive"
    msg_path = root / "Accounts" / "account@example.com" / "com.microsoft.__Messages" / "Inbox" / "message_00001.xml"
    msg_path.parent.mkdir(parents=True, exist_ok=True)
    msg_path.write_text(SAMPLE_XML, encoding="utf-8")

    attachments_dir = msg_path.parent / "com.microsoft.__Attachments"
    attachments_dir.mkdir()
    (attachments_dir / "invoice_0000").write_bytes(b"%PDF-1.7 fake")
    (attachments_dir / "meeting.ics").write_text("BEGIN:VCALENDAR\nEND:VCALENDAR\n", encoding="utf-8")

    parsed = parse_message_file(str(msg_path))

    assert parsed["message_id"] == "<abc123@example.com>"
    assert parsed["from_name"] == "Sender Name"
    assert parsed["from_addr"] == "sender@example.com"
    assert [item["addr"] for item in parsed["to"]] == ["alice@example.com", "bob@example.com"]
    assert [item["addr"] for item in parsed["cc"]] == ["carol@example.com"]
    assert parsed["subject"] == "Hello world"
    assert "2025" in parsed["date"]
    assert parsed["body_html"] is not None
    assert "Hello plain" in parsed["body_plain"]
    assert len(parsed["attachments"]) == 2
    assert parsed["attachments"][0]["filename"] == "invoice.pdf"
    assert parsed["attachments"][0]["data"] == b"%PDF-1.7 fake"


def test_parse_html_only_body_falls_back_to_plain_text(tmp_path: Path) -> None:
    xml = """<emails><email>
      <OPFMessageCopySenderAddress>
        <emailAddress OPFContactEmailAddressAddress="sender@example.com" OPFContactEmailAddressName="Sender" />
      </OPFMessageCopySenderAddress>
      <OPFMessageCopyBody>&lt;html&gt;&lt;body&gt;&lt;div&gt;Hi there&lt;/div&gt;&lt;/body&gt;&lt;/html&gt;</OPFMessageCopyBody>
    </email></emails>"""
    msg_path = tmp_path / "archive" / "Accounts" / "account@example.com" / "com.microsoft.__Messages" / "Inbox" / "message_00001.xml"
    msg_path.parent.mkdir(parents=True, exist_ok=True)
    msg_path.write_text(xml, encoding="utf-8")

    parsed = parse_message_file(str(msg_path))

    assert parsed["body_html"] is not None
    assert parsed["body_plain"] == "Hi there"
