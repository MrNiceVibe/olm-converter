import argparse
import shutil
import sys
from pathlib import Path

from rich.console import Console
from rich.table import Table

from .extractor import extract
from .writer import write_all


console = Console()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert Outlook for Mac OLM archives to MBOX/EML mail plus VCF/ICS supporting exports."
    )
    parser.add_argument(
        "source",
        help="Path to a .olm file or an already extracted Outlook archive directory.",
    )
    parser.add_argument("--out", default="./export", help="Output directory.")
    parser.add_argument(
        "--format",
        default="mbox,eml",
        help="Comma-separated mail formats: mbox, eml, or both. Contacts/calendars/tasks are exported automatically.",
    )
    args = parser.parse_args()

    source = Path(args.source).expanduser()
    if not source.exists():
        console.print(f"[red]Error:[/red] Source not found: {source}")
        sys.exit(1)

    formats = {item.strip().lower() for item in args.format.split(",") if item.strip()}
    invalid = formats - {"mbox", "eml"}
    if invalid or not formats:
        console.print(
            f"[red]Error:[/red] Unknown formats: {sorted(invalid) or args.format}. Use 'mbox', 'eml', or 'mbox,eml'."
        )
        sys.exit(1)

    console.print("[bold]OLM Converter[/bold]")
    console.print(f"  Source : {source}")
    console.print(f"  Output : {Path(args.out).resolve()}")
    console.print(f"  Format : {', '.join(sorted(formats))}")
    console.print("")

    try:
        result = extract(str(source))
    except (FileNotFoundError, ValueError) as exc:
        console.print(f"[red]Extraction failed:[/red] {exc}")
        sys.exit(1)

    console.print(
        f"Found [bold]{len(result.accounts)}[/bold] account(s) and [bold]{result.total_messages}[/bold] message file(s)."
    )

    try:
        summary = write_all(result, args.out, formats)
    finally:
        if result.temp_dir:
            shutil.rmtree(result.temp_dir, ignore_errors=True)

    table = Table(title="Conversion Summary")
    table.add_column("Metric")
    table.add_column("Count", justify="right")
    table.add_row("Messages written", str(summary["written"]))
    table.add_row("Messages skipped", str(summary["skipped"]))
    table.add_row("Contacts exported", str(summary["contacts_written"]))
    table.add_row("Calendar files exported", str(summary["calendar_files_written"]))
    table.add_row("Task files exported", str(summary["task_files_written"]))
    table.add_row("Errors", str(len(summary["errors"])))
    console.print("")
    console.print(table)

    if summary["errors"]:
        console.print("\n[yellow]Sample errors:[/yellow]")
        for entry in summary["errors"][:20]:
            console.print(f"  {entry['file']}: {entry['error']}")

    console.print(f"\n[green]Done.[/green] Output written to {Path(args.out).resolve()}")


if __name__ == "__main__":
    main()
