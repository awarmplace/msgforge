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

# Inline images — reference with cid: in HTML, attach with content_id
msg = Message(
    subject="Product Update",
    html_body="""
    <p>Hi team,</p>
    <p>Here's the new logo and sales chart:</p>
    <img src="cid:logo" width="200">
    <br>
    <img src="cid:chart" width="400">
    <p>Best regards</p>
    """,
    to=[("alice@example.com", "Alice")],
)
msg.attach("logo.png", content_id="logo")
msg.attach("chart.png", content_id="chart")
msg.attach("report.xlsx")  # regular attachment — shows in attachment bar
msg.save("product_update.msg")

# High importance
msg = Message(
    subject="Action Required",
    text_body="Please review ASAP.",
    to=[("bob@example.com", "Bob")],
    importance="high",
)
msg.save("urgent.msg")
```

## Features

- **HTML body** — full CSS, tables, bold/italic rendering in Outlook (via encapsulated HTML in RTF)
- **Plain text body** — fallback when no HTML provided
- **Inline images** — embed images in HTML via `cid:` references (hidden from attachment bar)
- **Recipients** — TO, CC, BCC with display names
- **File attachments** — appear in Outlook's attachment bar
- **Importance** — `"low"`, `"normal"`, or `"high"` (shown in Outlook's priority column)
- **Unicode support** — full Unicode in HTML bodies (smart quotes, CJK, emoji) via proper RTF Unicode escapes
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
    importance="normal",     # "low", "normal", or "high"
)
```

Recipients can be `("email", "Name")` tuples or just `"email"` strings.

### `.attach(path, filename=None, content_id=None)`

Attach a file from disk. Returns `self` for chaining. Pass `content_id` to embed as an inline image (`<img src="cid:content_id">`).

### `.attach_bytes(filename, data, mime_type=None, content_id=None)`

Attach a file from bytes. MIME type auto-detected from extension. Pass `content_id` for inline images.

### `.save(path)`

Write the `.msg` file to disk.

### `.as_bytes()`

Return the `.msg` file content as bytes (for web responses, etc).

## Limitations

- **Experimental** — implements MS-CFB and MS-OXMSG from scratch; not battle-tested across all Outlook versions
- **No read support** — this package only creates `.msg` files, it cannot read or parse them (use `extract-msg` for that)
- **No RTF body authoring** — rich text is generated from HTML only; direct RTF input is not supported
- **No digital signatures or encryption** — messages are unsigned and unencrypted
- **No calendar/contact/task items** — only email messages (`IPM.Note`) are supported
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
- Inline images use base64 data URIs embedded directly in the HTML before RTF encapsulation — Outlook's `\fromhtml1` renderer cannot resolve `cid:` references against the attachment table

## Running tests

```bash
pip install .[dev]
pytest tests/
```

## Built with AI

This package was built collaboratively with [Claude Code](https://claude.ai/claude-code) (Claude Opus 4.6). The OLE compound file writer, MAPI property encoding, and encapsulated HTML format were implemented from the MS-CFB and MS-OXMSG specifications through iterative development and testing.
