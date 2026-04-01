import tempfile
import zipfile
from pathlib import Path

from converter.extractor import extract


def _write_file(path: Path, content: str | bytes = "") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(content, bytes):
        path.write_bytes(content)
    else:
        path.write_text(content, encoding="utf-8")


def test_extract_from_directory_uses_real_messages_root(tmp_path: Path) -> None:
    root = tmp_path / "archive"
    _write_file(
        root / "Accounts" / "account@example.com" / "com.microsoft.__Messages" / "Inbox" / "message_00001.xml",
        "<emails><email /></emails>",
    )
    _write_file(
        root / "Accounts" / "account@example.com" / "com.microsoft.__Messages" / "Inbox" / "com.microsoft.__Attachments" / "file.pdf",
        b"pdf",
    )
    _write_file(root / "Accounts" / "account@example.com" / "Calendar" / "Calendar.xml", "<appointments />")

    result = extract(str(root))

    assert result.total_messages == 1
    assert list(result.accounts) == ["account@example.com"]
    inbox = result.accounts["account@example.com"].root_folders["Inbox"]
    assert len(inbox.message_files) == 1


def test_extract_from_olm_zip(tmp_path: Path) -> None:
    olm_path = tmp_path / "sample.olm"
    with zipfile.ZipFile(olm_path, "w") as zf:
        zf.writestr(
            "Accounts/account@example.com/com.microsoft.__Messages/Inbox/message_00001.xml",
            "<emails><email /></emails>",
        )

    result = extract(str(olm_path))

    assert result.total_messages == 1
    assert result.temp_dir is not None


def test_extract_handles_nested_folders(tmp_path: Path) -> None:
    root = tmp_path / "archive"
    _write_file(
        root
        / "Accounts"
        / "account@example.com"
        / "com.microsoft.__Messages"
        / "Projects"
        / "Client A"
        / "message_00001.xml",
        "<emails><email /></emails>",
    )

    result = extract(str(root))

    projects = result.accounts["account@example.com"].root_folders["Projects"]
    assert "Client A" in projects.subfolders
