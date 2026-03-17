# msgforge

Pure Python Outlook `.msg` file builder. No COM, no Outlook installation, no native dependencies. Works on Linux.

> **Experimental** — this package implements the MS-CFB and MS-OXMSG specs from scratch. It has been tested with Outlook on Windows but may not cover all edge cases. Use at your own risk.

## Install

```bash
pip install msgforge
```

## Usage

```python
from msgforge import Message

# Simple text email
msg = Message(
    subject="Hello",
    text_body="Hi there!",
    to=[("alice@example.com", "Alice")],
)
msg.save("hello.msg")

# HTML email with table, attachments, multiple recipients
msg = Message(
    subject="Weekly Report",
    html_body="""
    <p>Hi team,</p>
    <p>Please find the <b>weekly report</b> below:</p>
    <table border="1" style="border-collapse: collapse;">
        <tr><th style="padding: 4px 8px; background: #f0f0f0;">Region</th>
            <th style="padding: 4px 8px; background: #f0f0f0;">Revenue</th></tr>
        <tr><td style="padding: 4px 8px;">EMEA</td>
            <td style="padding: 4px 8px;">$1.2M</td></tr>
        <tr><td style="padding: 4px 8px;">APAC</td>
            <td style="padding: 4px 8px;">$800K</td></tr>
    </table>
    <p>Best regards</p>
    """,
    to=[("alice@example.com", "Alice"), ("bob@example.com", "Bob")],
    cc=[("carol@example.com", "Carol")],
)
msg.attach("report.xlsx")
msg.save("weekly_report.msg")
```

## Features

- **HTML body** — full CSS, tables, bold/italic rendering in Outlook (via encapsulated HTML in RTF)
- **Plain text body** — fallback when no HTML provided
- **Recipients** — TO, CC, BCC with display names
- **File attachments** — appear in Outlook's attachment bar
- **Pure Python** — only dependency is `compressed-rtf`
- **Cross-platform** — works on Linux, macOS, Windows

## API

### `Message(...)`

```python
Message(
    subject="...",           # Email subject
    html_body="<p>...</p>",  # HTML body (rendered in Outlook with full formatting)
    text_body="...",         # Plain text body (auto-generated from HTML if omitted)
    to=[("email", "Name")],  # TO recipients
    cc=[...],                # CC recipients
    bcc=[...],               # BCC recipients
)
```

Recipients can be `("email", "Name")` tuples or just `"email"` strings.

### `.attach(path, filename=None)`

Attach a file from disk. Returns `self` for chaining.

### `.attach_bytes(filename, data, mime_type=None)`

Attach a file from bytes. MIME type auto-detected from extension.

### `.save(path)`

Write the `.msg` file to disk.

### `.as_bytes()`

Return the `.msg` file content as bytes (for web responses, etc).

## Limitations

- **Experimental** — implements MS-CFB and MS-OXMSG from scratch; not battle-tested across all Outlook versions
- **No read support** — this package only creates `.msg` files, it cannot read or parse them (use `extract-msg` for that)
- **No embedded images** — inline images (`<img>` tags referencing CID attachments) are not supported
- **No RTF body authoring** — rich text is generated from HTML only; direct RTF input is not supported
- **No digital signatures or encryption** — messages are unsigned and unencrypted
- **No calendar/contact/task items** — only email messages (`IPM.Note`) are supported
- **ASCII-only RTF** — non-ASCII characters in HTML bodies are encoded as RTF escapes, which may not render correctly for all scripts
- **Recipient resolution** — recipients show display names but may not resolve to Exchange contacts until Outlook processes them
- **Large files** — no streaming support; the entire message is built in memory

## How it works

Implements two Microsoft specs from scratch:

1. **MS-CFB** (Compound File Binary Format) — the OLE structured storage container
2. **MS-OXMSG** (Outlook Item File Format) — MAPI properties, recipients, attachments

HTML bodies use the `\fromhtml1` encapsulated HTML format, the same format Outlook uses internally. This gives full HTML/CSS rendering including tables, styling, etc.

Key implementation details discovered through testing:
- Directory entry names must be sorted by **length first**, then case-insensitive (MS-CFB spec)
- `PT_UNICODE` property sizes must include +2 bytes for the null terminator
- HTML bodies require encapsulated HTML in RTF (`\fromhtml1`) — raw `PR_HTML` is not rendered by Outlook

## Running tests

```bash
pip install .[dev]
pytest tests/
```

## Built with AI

This package was built collaboratively with [Claude Code](https://claude.ai/claude-code) (Claude Opus 4.6). The OLE compound file writer, MAPI property encoding, and encapsulated HTML format were implemented from the MS-CFB and MS-OXMSG specifications through iterative development and testing.
