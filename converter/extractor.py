import os
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple


MESSAGE_EXTENSIONS = {".xml", ".olk15message"}
ATTACHMENTS_DIRNAME = "com.microsoft.__Attachments"
MESSAGES_DIRNAME = "com.microsoft.__Messages"
CONTACTS_FILENAME = "Contacts.xml"
CALENDAR_FILENAME = "Calendar.xml"
TASKS_FILENAME = "Tasks.xml"


@dataclass
class FolderNode:
    name: str
    path: str
    message_files: List[str] = field(default_factory=list)
    subfolders: Dict[str, "FolderNode"] = field(default_factory=dict)


@dataclass
class AccountNode:
    account_id: str
    display_name: str
    account_path: str
    messages_root: str
    root_folders: Dict[str, FolderNode] = field(default_factory=dict)
    contact_files: Dict[str, str] = field(default_factory=dict)
    calendar_files: Dict[str, str] = field(default_factory=dict)
    task_files: Dict[str, str] = field(default_factory=dict)


@dataclass
class ExtractionResult:
    source_root: str
    temp_dir: Optional[str]
    accounts: Dict[str, AccountNode]
    total_messages: int


def extract(source_path: str) -> ExtractionResult:
    """
    Accept either a raw .olm ZIP file or an already extracted archive folder.

    The real Outlook for Mac archive layout is:
    Accounts/<account>/com.microsoft.__Messages/.../message_*.xml
    """
    source = Path(source_path).expanduser()
    if not source.exists():
        raise FileNotFoundError(f"Source not found: {source_path}")

    source_root, accounts_root, temp_dir = _prepare_source(source)

    accounts: Dict[str, AccountNode] = {}
    total_messages = 0

    for account_dir in sorted(p for p in accounts_root.iterdir() if p.is_dir()):
        messages_root = account_dir / MESSAGES_DIRNAME

        account = AccountNode(
            account_id=account_dir.name,
            display_name=account_dir.name,
            account_path=str(account_dir),
            messages_root=str(messages_root),
        )

        if messages_root.is_dir():
            for dirpath, dirnames, filenames in os.walk(messages_root):
                dirnames[:] = sorted(
                    d for d in dirnames if d != ATTACHMENTS_DIRNAME and not d.startswith(".")
                )
                rel = Path(os.path.relpath(dirpath, messages_root))
                node = _get_or_create_folder(account.root_folders, rel.parts)

                for filename in sorted(filenames):
                    if _is_message_file(filename):
                        node.message_files.append(str(Path(dirpath) / filename))
                        total_messages += 1

        account.contact_files = _find_collection_files(account_dir, CONTACTS_FILENAME)
        account.calendar_files = _find_collection_files(account_dir, CALENDAR_FILENAME)
        account.task_files = _find_collection_files(account_dir, TASKS_FILENAME)

        if (
            not account.root_folders
            and not account.contact_files
            and not account.calendar_files
            and not account.task_files
        ):
            continue

        accounts[account.account_id] = account

    return ExtractionResult(
        source_root=str(source_root),
        temp_dir=temp_dir,
        accounts=accounts,
        total_messages=total_messages,
    )


def _prepare_source(source: Path) -> Tuple[Path, Path, Optional[str]]:
    if source.is_file():
        if not zipfile.is_zipfile(source):
            raise ValueError(f"Not a valid OLM/ZIP archive: {source}")
        temp_dir = tempfile.mkdtemp(prefix="olm_convert_")
        with zipfile.ZipFile(source, "r") as zf:
            zf.extractall(temp_dir)
        source_root = Path(temp_dir)
        accounts_root = source_root / "Accounts"
        if not accounts_root.is_dir():
            raise ValueError("Archive does not contain an Accounts directory.")
        return source_root, accounts_root, temp_dir

    if source.is_dir():
        if (source / "Accounts").is_dir():
            return source, source / "Accounts", None
        if source.name == "Accounts":
            return source.parent, source, None
        raise ValueError(
            "Directory source must be an extracted Outlook archive root or its Accounts directory."
        )

    raise ValueError(f"Unsupported source type: {source}")


def _get_or_create_folder(root: Dict[str, FolderNode], parts: Tuple[str, ...]) -> FolderNode:
    if not parts or parts == (".",):
        if "__root__" not in root:
            root["__root__"] = FolderNode(name="__root__", path=".")
        return root["__root__"]

    current = root
    current_parts: List[str] = []
    node: Optional[FolderNode] = None
    for part in parts:
        current_parts.append(part)
        if part not in current:
            current[part] = FolderNode(name=part, path=str(Path(*current_parts)))
        node = current[part]
        current = node.subfolders
    assert node is not None
    return node


def _is_message_file(filename: str) -> bool:
    path = Path(filename)
    return path.suffix.lower() in MESSAGE_EXTENSIONS and path.name.startswith("message_")


def _find_collection_files(account_dir: Path, filename: str) -> Dict[str, str]:
    files: Dict[str, str] = {}
    for path in sorted(account_dir.rglob(filename)):
        if any(part.startswith(".") for part in path.parts):
            continue
        rel_parent = path.parent.relative_to(account_dir)
        files[str(rel_parent)] = str(path)
    return files
