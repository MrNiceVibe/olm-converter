from email import message_from_bytes
from email.policy import default

from converter.builder import build_email_message


def make_message(**overrides):
    base = {
        "message_id": "<test@example.com>",
        "from_name": "Marcus",
        "from_addr": "marcus@example.com",
        "to": [{"name": "Alice", "addr": "alice@example.com"}],
        "cc": [],
        "bcc": [],
        "reply_to": [],
        "subject": "Test Subject",
        "date": "Sun, 01 Jun 2025 14:22:00 -0000",
        "body_plain": "Hello world",
        "body_html": "<html><body><p>Hello world</p></body></html>",
        "attachments": [
            {
                "filename": "doc.pdf",
                "content_type": "application/pdf",
                "content_id": "cid-1",
                "data": b"FAKEPDF",
            }
        ],
    }
    base.update(overrides)
    return base


def test_build_email_message_headers_and_body() -> None:
    em = build_email_message(make_message())

    assert em["Subject"] == "Test Subject"
    assert "marcus@example.com" in em["From"]
    assert "alice@example.com" in em["To"]
    assert em.is_multipart()


def test_build_email_message_serializes() -> None:
    em = build_email_message(make_message())
    parsed = message_from_bytes(em.as_bytes(), policy=default)

    assert parsed["Subject"] == "Test Subject"
    assert len(list(parsed.iter_attachments())) == 1
