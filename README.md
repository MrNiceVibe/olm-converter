# OLM Converter

`olm-converter` turns Outlook for Mac archives into portable email exports that are easier to inspect, back up, and import elsewhere.

Current outputs:

- `MBOX` for Gmail and mailbox migration workflows
- `EML` for one-message-per-file exports
- `VCF` for contacts
- `ICS` for standalone calendars
- `ICS` for tasks as `VTODO`

The parser is built around the real Outlook for Mac archive layout, including:

- extracted archive folders like `Accounts/<account>/com.microsoft.__Messages/...`
- message XML files such as `message_00001.xml`
- file-backed attachments referenced by `OPFMessageCopyAttachmentList`
- meeting sidecars referenced by `OPFMessageCopyMeetingData`

Notes are still ignored for now.

## Features

- Accepts either a raw `.olm` file or an already extracted archive folder
- Preserves account separation
- Preserves nested folder hierarchy
- Produces Gmail-friendly `.mbox` files
- Produces `.eml` files for direct message inspection
- Produces `.vcf` exports for Outlook contacts
- Produces `.ics` exports for Outlook calendars
- Produces `.ics` task lists with `VTODO` entries
- Includes tests and CI-friendly packaging

## Install

From source:

```bash
python -m pip install .
```

For development:

```bash
python -m pip install .[dev]
```

You can also use the local CLI without installing globally:

```bash
python -m converter.cli --help
```

## Usage

Convert a raw `.olm` file:

```bash
olm-converter /path/to/archive.olm --out ./export
```

Convert an already extracted Outlook archive directory:

```bash
olm-converter /path/to/extracted-archive --out ./export
```

Choose one output format:

```bash
olm-converter /path/to/archive.olm --out ./export --format mbox
olm-converter /path/to/archive.olm --out ./export --format eml
```

Or keep both:

```bash
olm-converter /path/to/archive.olm --out ./export --format mbox,eml
```

## Output Layout

```text
export/
├── contacts/
│   └── <account>/
│       └── <contact-folder>.vcf
├── calendars/
│   └── <account>/
│       └── <calendar-name>.ics
├── tasks/
│   └── <account>/
│       └── <task-list-name>.ics
├── eml/
│   └── <account>/
│       └── <folder tree>/
│           └── 00001.eml
└── mbox/
    └── <account>/
        └── <folder tree>.mbox
```

## Gmail Notes

For Gmail-oriented imports, `mbox` is the useful target format.

- One `.mbox` file usually maps best to one folder or label import workflow
- If you prefer a single mailbox file per account, you can merge folder-level MBOX files after export
- Very large imports may be easier through Gmail API-based tooling rather than the browser UI

## Mozilla / Thunderbird Notes

- Contacts can be imported from the generated `.vcf` files
- Calendars can be imported from the generated `.ics` files
- Tasks are exported as `.ics` files containing `VTODO` items

## Development

Run tests:

```bash
pytest
```

Run the CLI directly from source:

```bash
python -m converter.cli /path/to/archive.olm --out ./export
```

## Privacy

This repository is intended to be committed without private archive data. Keep personal `.olm` files, extracted archives, and generated exports out of version control.
