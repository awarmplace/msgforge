# Changelog

## 1.0.0 — 2026-07-06

First stable release.

### Fixed

- **Files over ~7 MB were silently corrupt.** The OLE writer only used the
  109 FAT-sector slots in the header and never emitted DIFAT sectors, so any
  message whose total size exceeded ~7 MB failed to open. DIFAT sectors are
  now written per MS-CFB; tested with attachments up to 60 MB.
- **Whitespace between adjacent HTML tags was dropped** during RTF
  encapsulation, so `<b>bold</b> <i>italic</i>` rendered as "bolditalic".
  Whitespace-only text runs now survive as a single space, and newlines
  inside text (which RTF readers ignore) are converted to spaces.
- Attaching a file without an extension no longer writes a zero-length
  `PidTagAttachExtension` stream (forbidden by MS-OXMSG).
- An invalid `importance` value now raises `ValueError` instead of being
  silently treated as `"normal"`.
- `attach(path, filename=...)` now guesses the MIME type from the overridden
  filename instead of the on-disk name.
- `<style>`/`<script>` contents and HTML comments no longer leak into the
  auto-generated plain-text body.
- Recipients with an empty display name fall back to the email address
  (avoids zero-length property streams).

### Added

- `sender=` — set the From address (`PidTagSenderName`/`EmailAddress`/
  `SmtpAddress` plus the SentRepresenting mirrors).
- `sent=` — mark a message as sent/received instead of an unsent draft.
  Pass `True` or a `datetime` (also sets `PidTagClientSubmitTime` and
  `PidTagMessageDeliveryTime`).
- `PidTagHtml` (0x1013) is now written alongside the encapsulated RTF, so
  non-Outlook consumers (extract-msg, converters, e-discovery tools) get the
  raw HTML directly.
- Accurate type hints on the public API (`Optional`/`Sequence`, `Recipient`
  alias) to match the shipped `py.typed`.
- GitHub Actions CI (pytest on Python 3.9–3.13 × Linux/Windows/macOS, ruff).

### Changed

- Output is now byte-for-byte deterministic: directory-entry timestamps are
  zero (as MS-CFB recommends for the root storage) instead of the current
  time.

## 0.2.1

- MS-CFB/MS-OXMSG spec compliance fixes.

## 0.2.0

- Inline images (`cid:`), importance, Unicode fix, code cleanup.

## 0.1.1

- Initial release.
