import mailbox
from pathlib import Path
from typing import Iterable, Optional, Set

from rich.console import Console
from rich.progress import BarColumn, Progress, SpinnerColumn, TaskProgressColumn, TextColumn

from .ancillary import (
    parse_calendar_file,
    parse_contacts_file,
    parse_tasks_file,
    render_calendar_ics,
    render_tasks_ics,
    render_vcf,
)
from .builder import build_email_message
from .extractor import AccountNode, ExtractionResult, FolderNode
from .parser import parse_message_file


console = Console()


def write_all(result: ExtractionResult, out_dir: str, formats: Set[str]) -> dict:
    out = Path(out_dir)
    summary = {
        "written": 0,
        "skipped": 0,
        "errors": [],
        "contacts_written": 0,
        "calendar_files_written": 0,
        "task_files_written": 0,
    }

    if "mbox" in formats:
        (out / "mbox").mkdir(parents=True, exist_ok=True)
    if "eml" in formats:
        (out / "eml").mkdir(parents=True, exist_ok=True)
    (out / "contacts").mkdir(parents=True, exist_ok=True)
    (out / "calendars").mkdir(parents=True, exist_ok=True)
    (out / "tasks").mkdir(parents=True, exist_ok=True)

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        console=console,
        transient=True,
    ) as progress:
        task = progress.add_task("Converting...", total=result.total_messages)
        for account in result.accounts.values():
            _process_account(account, out, formats, summary, progress, task)
            _write_account_supporting_data(account, out, summary)

    return summary


def _process_account(
    account: AccountNode,
    out_dir: Path,
    formats: Set[str],
    summary: dict,
    progress: Progress,
    task: int,
) -> None:
    account_name = _safe_name(account.display_name)
    for folder in account.root_folders.values():
        _process_folder(folder, account_name, [], out_dir, formats, summary, progress, task)


def _process_folder(
    folder: FolderNode,
    account_name: str,
    path_parts: list[str],
    out_dir: Path,
    formats: Set[str],
    summary: dict,
    progress: Progress,
    task: int,
) -> None:
    current_parts = path_parts if folder.name == "__root__" else path_parts + [folder.name]

    if folder.message_files:
        _write_folder_messages(
            files=folder.message_files,
            account_name=account_name,
            path_parts=current_parts,
            out_dir=out_dir,
            formats=formats,
            summary=summary,
            progress=progress,
            task=task,
        )

    for subfolder in folder.subfolders.values():
        _process_folder(subfolder, account_name, current_parts, out_dir, formats, summary, progress, task)


def _write_folder_messages(
    files: Iterable[str],
    account_name: str,
    path_parts: list[str],
    out_dir: Path,
    formats: Set[str],
    summary: dict,
    progress: Progress,
    task: int,
) -> None:
    folder_path = Path(*path_parts) if path_parts else Path("_root")

    mbox_obj: Optional[mailbox.mbox] = None
    if "mbox" in formats:
        mbox_dir = out_dir / "mbox" / account_name / folder_path.parent
        mbox_dir.mkdir(parents=True, exist_ok=True)
        mbox_path = mbox_dir / f"{_safe_name(folder_path.name)}.mbox"
        mbox_obj = mailbox.mbox(str(mbox_path))
        mbox_obj.lock()

    eml_dir: Optional[Path] = None
    if "eml" in formats:
        eml_dir = out_dir / "eml" / account_name / folder_path
        eml_dir.mkdir(parents=True, exist_ok=True)

    try:
        for index, filepath in enumerate(sorted(files), start=1):
            try:
                msg_dict = parse_message_file(filepath)
                email_message = build_email_message(msg_dict)

                if mbox_obj is not None:
                    mbox_obj.add(mailbox.mboxMessage(email_message))

                if eml_dir is not None:
                    eml_path = eml_dir / f"{index:05d}.eml"
                    eml_path.write_bytes(email_message.as_bytes())

                summary["written"] += 1
            except Exception as exc:
                summary["errors"].append({"file": filepath, "error": str(exc)})
                summary["skipped"] += 1

            progress.advance(task)
    finally:
        if mbox_obj is not None:
            mbox_obj.flush()
            mbox_obj.unlock()
            mbox_obj.close()


def _safe_name(value: str) -> str:
    value = value.strip().replace("/", "_")
    return "".join(char if char not in {"\0", ":"} else "_" for char in value) or "_"


def _write_account_supporting_data(account: AccountNode, out_dir: Path, summary: dict) -> None:
    account_name = _safe_name(account.display_name)
    _write_contacts(account_name, account.contact_files, out_dir, summary)
    _write_calendars(account_name, account.calendar_files, out_dir, summary)
    _write_tasks(account_name, account.task_files, out_dir, summary)


def _write_contacts(account_name: str, files: dict[str, str], out_dir: Path, summary: dict) -> None:
    for rel_parent, filepath in files.items():
        try:
            contacts = parse_contacts_file(filepath)
            if not contacts:
                continue
            out_path = _collection_output_path(out_dir / "contacts" / account_name, rel_parent, ".vcf")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(render_vcf(contacts), encoding="utf-8")
            summary["contacts_written"] += len(contacts)
        except Exception as exc:
            summary["errors"].append({"file": filepath, "error": str(exc)})


def _write_calendars(account_name: str, files: dict[str, str], out_dir: Path, summary: dict) -> None:
    for rel_parent, filepath in files.items():
        try:
            calendar = parse_calendar_file(filepath)
            out_path = _collection_output_path(out_dir / "calendars" / account_name, rel_parent, ".ics")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(render_calendar_ics(calendar, Path(rel_parent).name), encoding="utf-8")
            summary["calendar_files_written"] += 1
        except Exception as exc:
            summary["errors"].append({"file": filepath, "error": str(exc)})


def _write_tasks(account_name: str, files: dict[str, str], out_dir: Path, summary: dict) -> None:
    for rel_parent, filepath in files.items():
        try:
            tasks = parse_tasks_file(filepath)
            out_path = _collection_output_path(out_dir / "tasks" / account_name, rel_parent, ".ics")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(render_tasks_ics(tasks, Path(rel_parent).name), encoding="utf-8")
            summary["task_files_written"] += 1
        except Exception as exc:
            summary["errors"].append({"file": filepath, "error": str(exc)})


def _collection_output_path(base_dir: Path, rel_parent: str, suffix: str) -> Path:
    rel_path = Path(rel_parent)
    parts = rel_path.parts
    if not parts:
        return base_dir / f"root{suffix}"
    target_dir = base_dir / Path(*parts[:-1]) if len(parts) > 1 else base_dir
    return target_dir / f"{_safe_name(parts[-1])}{suffix}"
