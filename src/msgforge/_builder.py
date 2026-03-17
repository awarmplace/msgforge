"""msgforge — Pure Python Outlook .msg file builder.

Create Outlook .msg files with HTML bodies, recipients, and attachments.
No COM, no Outlook installation, no native dependencies. Works on Linux.

Usage::

    from msgforge import Message

    msg = Message(
        subject="Q1 Report",
        html_body="<p>Please see <b>attached</b> report.</p>",
        to=[("alice@example.com", "Alice"), ("bob@example.com", "Bob")],
        cc=[("carol@example.com", "Carol")],
    )
    msg.attach("report.xlsx")
    msg.save("Q1 Report.msg")

Or build incrementally::

    msg = Message()
    msg.subject = "Hello"
    msg.html_body = "<p>Hi there</p>"
    msg.to = [("user@example.com", "User Name")]
    msg.attach_bytes("data.csv", b"a,b,c\\n1,2,3")
    msg.save("hello.msg")

Requires: compressed-rtf (pip install compressed-rtf)
"""

from __future__ import annotations

import re
import struct
import time
import html as html_mod
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Tuple, Union


# ─── Public API ──────────────────────────────────────────────────────────────

class Message:
    """An Outlook .msg email message.

    Args:
        subject: Email subject line.
        html_body: HTML body (rendered with full CSS/table support in Outlook).
        text_body: Plain text body (used as fallback, or auto-generated from HTML).
        to: List of TO recipients as ``("email", "Display Name")`` tuples.
            The display name is optional: ``"email"`` alone also works.
        cc: List of CC recipients (same format as ``to``).
        bcc: List of BCC recipients (same format as ``to``).
    """

    def __init__(
        self,
        subject: str = "",
        html_body: str = "",
        text_body: str = "",
        to: list = None,
        cc: list = None,
        bcc: list = None,
    ):
        self.subject = subject
        self.html_body = html_body
        self.text_body = text_body
        self.to: List[Tuple[str, str]] = _normalize_recipients(to)
        self.cc: List[Tuple[str, str]] = _normalize_recipients(cc)
        self.bcc: List[Tuple[str, str]] = _normalize_recipients(bcc)
        self._attachments: List[Tuple[str, bytes, str]] = []

    def attach(self, path: Union[str, Path], filename: str = None) -> Message:
        """Attach a file from disk.

        Args:
            path: Path to the file.
            filename: Override the filename shown in Outlook.

        Returns:
            self (for chaining).
        """
        path = Path(path)
        self._attachments.append((
            filename or path.name,
            path.read_bytes(),
            _guess_mime(path.name),
        ))
        return self

    def attach_bytes(self, filename: str, data: bytes,
                     mime_type: str = None) -> Message:
        """Attach a file from bytes.

        Args:
            filename: Filename shown in Outlook.
            data: Raw file content.
            mime_type: MIME type (auto-detected from extension if omitted).

        Returns:
            self (for chaining).
        """
        self._attachments.append((filename, data, mime_type or _guess_mime(filename)))
        return self

    def save(self, path: Union[str, Path]) -> None:
        """Save the message as an Outlook .msg file."""
        _build_msg(self, path)

    def as_bytes(self) -> bytes:
        """Return the .msg file content as bytes."""
        import io
        buf = io.BytesIO()
        # Write to a temp path then read back
        import tempfile
        with tempfile.NamedTemporaryFile(suffix='.msg', delete=False) as f:
            tmp = f.name
        try:
            _build_msg(self, tmp)
            return Path(tmp).read_bytes()
        finally:
            Path(tmp).unlink(missing_ok=True)


# ─── Recipient helpers ───────────────────────────────────────────────────────

def _normalize_recipients(recipients) -> List[Tuple[str, str]]:
    """Normalize recipient list to [(email, display_name), ...]."""
    if not recipients:
        return []
    result = []
    for r in recipients:
        if isinstance(r, str):
            result.append((r, r))
        elif isinstance(r, (tuple, list)) and len(r) >= 2:
            result.append((r[0], r[1]))
        elif isinstance(r, (tuple, list)) and len(r) == 1:
            result.append((r[0], r[0]))
        else:
            result.append((str(r), str(r)))
    return result


# ─── OLE Compound File constants (MS-CFB) ───────────────────────────────────

_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
_SECTOR_SIZE = 512
_MINI_SECTOR_SIZE = 64
_MINI_STREAM_CUTOFF = 0x1000
_FREESECT = 0xFFFFFFFF
_ENDOFCHAIN = 0xFFFFFFFE
_FATSECT = 0xFFFFFFFD
_NOSTREAM = 0xFFFFFFFF
_DIR_ENTRY_SIZE = 128
_ENTRIES_PER_FAT_SECTOR = _SECTOR_SIZE // 4

_STGTY_STORAGE = 1
_STGTY_STREAM = 2
_STGTY_ROOT = 5
_RED = 0
_BLACK = 1

_MSG_CLSID = bytes([
    0x0B, 0x0D, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
])

# ─── MAPI constants ─────────────────────────────────────────────────────────

_PT_LONG = 0x0003
_PT_BOOLEAN = 0x000B
_PT_UNICODE = 0x001F
_PT_BINARY = 0x0102
_MAPI_TO = 1
_MAPI_CC = 2
_MAPI_BCC = 3


# ─── OLE Writer (internal) ──────────────────────────────────────────────────

@dataclass
class _DirEntry:
    name: str
    entry_type: int
    data: bytes = b''
    children: list = field(default_factory=list)
    clsid: bytes = b'\x00' * 16
    _dir_id: int = 0
    _left_id: int = _NOSTREAM
    _right_id: int = _NOSTREAM
    _child_id: int = _NOSTREAM
    _color: int = _RED
    _start_sector: int = _ENDOFCHAIN
    _data_size: int = 0


class _OleWriter:
    """Writes OLE Compound Binary Files (MS-CFB v3, 512-byte sectors)."""

    def __init__(self):
        self._root = _DirEntry("Root Entry", _STGTY_ROOT)

    @property
    def root(self) -> _DirEntry:
        return self._root

    def add_storage(self, parent: _DirEntry, name: str) -> _DirEntry:
        entry = _DirEntry(name, _STGTY_STORAGE)
        parent.children.append(entry)
        return entry

    def add_stream(self, parent: _DirEntry, name: str, data: bytes) -> _DirEntry:
        entry = _DirEntry(name, _STGTY_STREAM, data)
        entry._data_size = len(data)
        parent.children.append(entry)
        return entry

    def write(self, path: Union[str, Path]) -> None:
        entries = self._flatten_sorted()

        large = [e for e in entries if e.entry_type == _STGTY_STREAM and e.data
                 and len(e.data) >= _MINI_STREAM_CUTOFF]
        small = [e for e in entries if e.entry_type == _STGTY_STREAM and e.data
                 and len(e.data) < _MINI_STREAM_CUTOFF]

        mini_size = sum((len(e.data) + 63) >> 6 for e in small)
        fat_size = sum((len(e.data) + 511) >> 9 for e in large)
        dir_cnt = (len(entries) + 3) >> 2
        mfat_cnt = (mini_size + 127) >> 7
        mini_container_cnt = (mini_size + 7) >> 3

        fat_base = fat_size + dir_cnt + mfat_cnt + mini_container_cnt
        fat_cnt = 0
        while True:
            needed = (fat_base + fat_cnt + _ENTRIES_PER_FAT_SECTOR - 1) >> 7
            if needed <= fat_cnt:
                break
            fat_cnt = needed

        total_sectors = fat_cnt + mfat_cnt + dir_cnt + fat_size + mini_container_cnt

        # Build FAT
        fat = [_FATSECT] * fat_cnt
        i = fat_cnt

        def chainit(count):
            nonlocal i
            for _ in range(count - 1):
                fat.append(i + 1)
                i += 1
            if count:
                fat.append(_ENDOFCHAIN)
                i += 1

        chainit(mfat_cnt)
        dir_start = i
        chainit(dir_cnt)

        for e in large:
            e._start_sector = i
            chainit((len(e.data) + 511) >> 9)

        mini_container_start = i if mini_container_cnt > 0 else _ENDOFCHAIN
        chainit(mini_container_cnt)

        self._root._start_sector = mini_container_start if mini_size > 0 else _ENDOFCHAIN
        self._root._data_size = mini_size * _MINI_SECTOR_SIZE if mini_size > 0 else 0

        # Mini-FAT
        mini_fat = []
        mi = 0
        for e in small:
            if not e.data:
                continue
            e._start_sector = mi
            n_mini = (len(e.data) + 63) >> 6
            for k in range(n_mini - 1):
                mini_fat.append(mi + 1)
                mi += 1
            mini_fat.append(_ENDOFCHAIN)
            mi += 1

        while len(fat) < fat_cnt * _ENTRIES_PER_FAT_SECTOR:
            fat.append(_FREESECT)

        # Build output
        buf = bytearray(_SECTOR_SIZE * (1 + total_sectors))
        self._write_header(buf, fat_cnt, dir_start,
                           fat_cnt if mfat_cnt > 0 else _ENDOFCHAIN, mfat_cnt)

        for si in range(fat_cnt):
            off = _SECTOR_SIZE * (1 + si)
            for j in range(_ENTRIES_PER_FAT_SECTOR):
                struct.pack_into('<I', buf, off + j * 4,
                                fat[si * _ENTRIES_PER_FAT_SECTOR + j] & 0xFFFFFFFF)

        mfat_off = _SECTOR_SIZE * (1 + fat_cnt)
        for j, v in enumerate(mini_fat):
            struct.pack_into('<I', buf, mfat_off + j * 4, v & 0xFFFFFFFF)
        for j in range(len(mini_fat), mfat_cnt * _ENTRIES_PER_FAT_SECTOR):
            struct.pack_into('<I', buf, mfat_off + j * 4, _FREESECT)

        dir_off = _SECTOR_SIZE * (1 + fat_cnt + mfat_cnt)
        for idx, e in enumerate(entries):
            self._write_dir_entry(buf, dir_off + idx * _DIR_ENTRY_SIZE, e)
        for idx in range(len(entries), dir_cnt * 4):
            off = dir_off + idx * _DIR_ENTRY_SIZE
            struct.pack_into('<I', buf, off + 68, _NOSTREAM)
            struct.pack_into('<I', buf, off + 72, _NOSTREAM)
            struct.pack_into('<I', buf, off + 76, _NOSTREAM)

        large_off = _SECTOR_SIZE * (1 + fat_cnt + mfat_cnt + dir_cnt)
        for e in large:
            eoff = large_off + (e._start_sector - fat_cnt - mfat_cnt - dir_cnt) * _SECTOR_SIZE
            buf[eoff:eoff + len(e.data)] = e.data

        mini_off = _SECTOR_SIZE * (1 + fat_cnt + mfat_cnt + dir_cnt + fat_size)
        for e in small:
            if not e.data:
                continue
            soff = mini_off + e._start_sector * _MINI_SECTOR_SIZE
            buf[soff:soff + len(e.data)] = e.data

        with open(path, 'wb') as f:
            f.write(buf)

    def _flatten_sorted(self) -> List[_DirEntry]:
        pairs: List[Tuple[str, _DirEntry]] = []

        def visit(entry: _DirEntry, path: str):
            pairs.append((path, entry))
            sorted_children = sorted(entry.children,
                                     key=lambda e: (len(e.name), e.name.upper()))
            for child in sorted_children:
                child_path = path + child.name
                if child.entry_type in (_STGTY_STORAGE, _STGTY_ROOT):
                    child_path += "/"
                visit(child, child_path)

        visit(self._root, "/")
        root_pair = pairs[0]
        rest = sorted(pairs[1:], key=lambda p: (len(p[0]), p[0].upper()))
        sorted_pairs = [root_pair] + rest

        for idx, (p, e) in enumerate(sorted_pairs):
            e._dir_id = idx
            e._left_id = _NOSTREAM
            e._right_id = _NOSTREAM
            e._child_id = _NOSTREAM
            e._color = _BLACK

        entries = [e for _, e in sorted_pairs]
        paths = [p for p, _ in sorted_pairs]

        def get_parent(p: str) -> str:
            if p.endswith('/'):
                p = p.rstrip('/')
            idx = p.rfind('/')
            return '/' if idx <= 0 else p[:idx + 1]

        from collections import defaultdict
        children_of: dict = defaultdict(list)
        for idx in range(1, len(entries)):
            children_of[get_parent(paths[idx])].append(idx)

        def build_bst(indices: list, depth: int = 0, max_black: int = -1) -> int:
            if not indices:
                return _NOSTREAM
            if max_black < 0:
                max_black = len(indices).bit_length() - 1 if indices else 0
            mid = len(indices) // 2
            root_idx = indices[mid]
            entries[root_idx]._color = _RED if depth >= max_black else _BLACK
            entries[root_idx]._left_id = build_bst(indices[:mid], depth + 1, max_black)
            entries[root_idx]._right_id = build_bst(indices[mid + 1:], depth + 1, max_black)
            return root_idx

        root_kids = children_of.get('/', [])
        if root_kids:
            entries[0]._child_id = build_bst(root_kids)

        for idx in range(1, len(entries)):
            if entries[idx].entry_type == _STGTY_STORAGE:
                storage_path = paths[idx] if paths[idx].endswith('/') else paths[idx] + '/'
                kids = children_of.get(storage_path, [])
                if kids:
                    entries[idx]._child_id = build_bst(kids)

        return entries

    def _write_dir_entry(self, buf: bytearray, off: int, entry: _DirEntry) -> None:
        name_bytes = entry.name.encode('utf-16-le')[:62]
        buf[off:off + len(name_bytes)] = name_bytes
        buf[off + len(name_bytes)] = 0
        buf[off + len(name_bytes) + 1] = 0
        struct.pack_into('<H', buf, off + 64, len(name_bytes) + 2)
        buf[off + 66] = entry.entry_type
        buf[off + 67] = _RED if entry.entry_type == _STGTY_ROOT else entry._color
        struct.pack_into('<I', buf, off + 68, entry._left_id & 0xFFFFFFFF)
        struct.pack_into('<I', buf, off + 72, entry._right_id & 0xFFFFFFFF)
        struct.pack_into('<I', buf, off + 76, entry._child_id & 0xFFFFFFFF)
        buf[off + 80:off + 96] = entry.clsid
        if entry.entry_type in (_STGTY_STORAGE, _STGTY_ROOT):
            ft = _filetime_now()
            struct.pack_into('<Q', buf, off + 100, ft)
            struct.pack_into('<Q', buf, off + 108, ft)
        if entry.entry_type == _STGTY_ROOT:
            struct.pack_into('<I', buf, off + 116, entry._start_sector & 0xFFFFFFFF)
        elif entry.entry_type == _STGTY_STREAM:
            struct.pack_into('<I', buf, off + 116, entry._start_sector & 0xFFFFFFFF)
        struct.pack_into('<I', buf, off + 120, entry._data_size)

    def _write_header(self, buf, fat_cnt, dir_start, mfat_start, mfat_cnt):
        buf[0:8] = _MAGIC
        struct.pack_into('<H', buf, 24, 0x003E)
        struct.pack_into('<H', buf, 26, 0x0003)
        struct.pack_into('<H', buf, 28, 0xFFFE)
        struct.pack_into('<H', buf, 30, 9)
        struct.pack_into('<H', buf, 32, 6)
        struct.pack_into('<I', buf, 44, fat_cnt)
        struct.pack_into('<I', buf, 48, dir_start)
        struct.pack_into('<I', buf, 56, _MINI_STREAM_CUTOFF)
        struct.pack_into('<I', buf, 60, mfat_start if mfat_cnt > 0 else _ENDOFCHAIN)
        struct.pack_into('<I', buf, 64, mfat_cnt)
        struct.pack_into('<I', buf, 68, _ENDOFCHAIN)
        struct.pack_into('<I', buf, 72, 0)
        for i in range(109):
            struct.pack_into('<I', buf, 76 + i * 4, i if i < fat_cnt else _FREESECT)


# ─── MAPI Property helpers (internal) ───────────────────────────────────────

class _Props:
    def __init__(self):
        self._entries: List[Tuple[int, int, int]] = []

    def add_long(self, prop_id, value):
        self._entries.append((prop_id, _PT_LONG, value))

    def add_boolean(self, prop_id, value):
        self._entries.append((prop_id, _PT_BOOLEAN, 1 if value else 0))

    def add_unicode(self, ole, parent, prop_id, value):
        data = value.encode('utf-16-le')
        ole.add_stream(parent, f"__substg1.0_{prop_id:04X}{_PT_UNICODE:04X}", data)
        self._entries.append((prop_id, _PT_UNICODE, len(data) + 2))

    def add_binary(self, ole, parent, prop_id, value):
        ole.add_stream(parent, f"__substg1.0_{prop_id:04X}{_PT_BINARY:04X}", value)
        self._entries.append((prop_id, _PT_BINARY, len(value)))

    def build_msg_stream(self, num_recip, num_attach):
        header = bytearray(32)
        struct.pack_into('<I', header, 8, num_recip)
        struct.pack_into('<I', header, 12, num_attach)
        struct.pack_into('<I', header, 16, num_recip)
        struct.pack_into('<I', header, 20, num_attach)
        return bytes(header) + self._build_entries()

    def build_sub_stream(self):
        return b'\x00' * 8 + self._build_entries()

    def _build_entries(self):
        buf = bytearray()
        for prop_id, prop_type, value in self._entries:
            entry = bytearray(16)
            struct.pack_into('<H', entry, 0, prop_type)
            struct.pack_into('<H', entry, 2, prop_id)
            struct.pack_into('<I', entry, 4, 0x00000006)
            if prop_type in (_PT_LONG, _PT_BOOLEAN):
                struct.pack_into('<I', entry, 8, value & 0xFFFFFFFF)
            elif prop_type in (_PT_UNICODE, _PT_BINARY):
                struct.pack_into('<I', entry, 8, value)
            buf.extend(entry)
        return bytes(buf)


# ─── MSG builder (internal) ─────────────────────────────────────────────────

def _build_msg(msg: Message, path: Union[str, Path]) -> None:
    """Build the .msg file from a Message object."""
    ole = _OleWriter()
    root = ole.root
    root.clsid = _MSG_CLSID

    # Named property mapping (required by spec, empty is fine)
    nameid = ole.add_storage(root, "__nameid_version1.0")
    ole.add_stream(nameid, "__substg1.0_00020102", b'')
    ole.add_stream(nameid, "__substg1.0_00030102", b'')
    ole.add_stream(nameid, "__substg1.0_00040102", b'')

    # Collect all recipients
    all_recip = ([(e, n, _MAPI_TO) for e, n in msg.to]
                 + [(e, n, _MAPI_CC) for e, n in msg.cc]
                 + [(e, n, _MAPI_BCC) for e, n in msg.bcc])

    # Message properties
    mp = _Props()
    mp.add_unicode(ole, root, 0x001A, "IPM.Note")       # MessageClass
    mp.add_unicode(ole, root, 0x0037, msg.subject or "") # Subject
    mp.add_unicode(ole, root, 0x003D, "")                # SubjectPrefix
    mp.add_unicode(ole, root, 0x0070, msg.subject or "") # ConversationTopic

    if msg.html_body:
        import compressed_rtf
        rtf_bytes = _encapsulate_html(msg.html_body)
        mp.add_binary(ole, root, 0x1009, compressed_rtf.compress(rtf_bytes))
        mp.add_boolean(0x0E1F, True)  # RTF_IN_SYNC
        mp.add_unicode(ole, root, 0x1000,
                       msg.text_body or _strip_html(msg.html_body))
    elif msg.text_body:
        mp.add_unicode(ole, root, 0x1000, msg.text_body)

    # Display headers (always include, even if empty)
    mp.add_unicode(ole, root, 0x0E04,
                   "; ".join(n for _, n, t in all_recip if t == _MAPI_TO))
    mp.add_unicode(ole, root, 0x0E03,
                   "; ".join(n for _, n, t in all_recip if t == _MAPI_CC))
    mp.add_unicode(ole, root, 0x0E02,
                   "; ".join(n for _, n, t in all_recip if t == _MAPI_BCC))

    flags = 0x08  # MSGFLAG_UNSENT
    if msg._attachments:
        flags |= 0x10  # MSGFLAG_HASATTACH
    mp.add_long(0x0E07, flags)           # MessageFlags
    mp.add_long(0x340D, 0x00040000)      # StoreSupportMask: STORE_UNICODE_OK
    mp.add_long(0x3FDE, 65001)           # InternetCodepage: UTF-8
    mp.add_long(0x3FF1, 0x0409)          # MessageLocaleId: en-US
    mp.add_long(0x0017, 1)               # Importance: normal
    mp.add_long(0x0026, 0)               # Priority: none
    mp.add_long(0x0036, 0)               # Sensitivity: none
    if msg._attachments:
        mp.add_boolean(0x0E1B, True)     # HasAttach

    ole.add_stream(root, "__properties_version1.0",
                   mp.build_msg_stream(len(all_recip), len(msg._attachments)))

    # Recipients
    for i, (email, display_name, recip_type) in enumerate(all_recip):
        stor = ole.add_storage(root, f"__recip_version1.0_#{i:08X}")
        rp = _Props()
        rp.add_unicode(ole, stor, 0x3001, display_name)   # DisplayName
        rp.add_unicode(ole, stor, 0x3002, "SMTP")         # AddrType
        rp.add_unicode(ole, stor, 0x3003, email)           # EmailAddress
        rp.add_unicode(ole, stor, 0x39FE, email)           # SmtpAddress
        rp.add_unicode(ole, stor, 0x5FF6, display_name)    # RecipientDisplayName
        rp.add_long(0x0C15, recip_type)   # RecipientType
        rp.add_long(0x0FFE, 6)            # ObjectType: MAPI_MAILUSER
        rp.add_long(0x3900, 0)            # DisplayType: DT_MAILUSER
        rp.add_long(0x3000, i)            # RowId
        rp.add_long(0x5FDF, 0)            # RecipientFlags
        rp.add_long(0x5FFD, 1)            # RecipientTrackStatus
        rp.add_long(0x5FFF, 0)            # RecipientOrder
        ole.add_stream(stor, "__properties_version1.0", rp.build_sub_stream())

    # Attachments
    for i, (filename, data, mime_type) in enumerate(msg._attachments):
        stor = ole.add_storage(root, f"__attach_version1.0_#{i:08X}")
        ap = _Props()
        ext = Path(filename).suffix
        short_name = filename[:8] + ext if len(filename) > 12 else filename
        ap.add_long(0x0E21, i)                  # AttachNumber
        ap.add_long(0x0FFE, 7)                  # ObjectType: MAPI_ATTACH
        ap.add_long(0x3705, 1)                  # AttachMethod: BY_VALUE
        ap.add_long(0x370B, 0xFFFFFFFF)         # RenderingPosition: not inline
        ap.add_long(0x3714, 0)                  # AttachFlags: normal
        ap.add_unicode(ole, stor, 0x3704, short_name)   # AttachFilename
        ap.add_unicode(ole, stor, 0x3707, filename)     # AttachLongFilename
        ap.add_unicode(ole, stor, 0x3703, ext)          # AttachExtension
        ap.add_unicode(ole, stor, 0x3001, filename)     # DisplayName
        ap.add_binary(ole, stor, 0x3701, data)          # AttachDataBinary
        if mime_type:
            ap.add_unicode(ole, stor, 0x370E, mime_type)  # AttachMimeTag
        ole.add_stream(stor, "__properties_version1.0", ap.build_sub_stream())

    ole.write(path)


# ─── HTML / RTF helpers (internal) ──────────────────────────────────────────

def _encapsulate_html(html_str: str) -> bytes:
    """Encapsulate HTML in RTF using the \\fromhtml1 format.

    This is the standard format Outlook uses internally: HTML tags are
    wrapped in {\\*\\htmltag} groups inside an RTF document. Outlook
    renders it as full HTML with CSS, tables, etc.
    """
    _ = chr(92)
    out = ['{' + _ + 'rtf1' + _ + 'ansi' + _ + 'ansicpg1252' + _ + 'fromhtml1 '
           + _ + 'deff0{' + _ + 'fonttbl{' + _ + 'f0' + _ + 'fswiss Arial;}}'
           + _ + 'uc1' + _ + 'pard' + _ + 'plain' + _ + 'deftab360 '
           + _ + 'f0' + _ + 'fs24 ']

    for part in re.split(r'(<[^>]+>)', html_str):
        if not part:
            continue
        encoded = _rtf_encode(part)
        if part.startswith('<'):
            out.append('{' + _ + '*' + _ + 'htmltag ' + encoded + '}')
        elif encoded.strip():
            out.append(encoded)

    out.append('}')
    return ''.join(out).encode('ascii', errors='replace')


def _rtf_encode(text: str) -> str:
    """Encode text for RTF, escaping special chars and non-ASCII."""
    _ = chr(92)
    out = []
    for ch in text:
        if ch == '\\':
            out.append(_ + _)
        elif ch == '{':
            out.append(_ + '{')
        elif ch == '}':
            out.append(_ + '}')
        elif ord(ch) > 127:
            out.append(_ + "'" + format(ord(ch) & 0xFF, '02x'))
        else:
            out.append(ch)
    return ''.join(out)


def _strip_html(html_str: str) -> str:
    """Strip HTML tags for plain text fallback."""
    text = re.sub(r'<br\s*/?>', '\n', html_str)
    text = re.sub(r'</p>', '\n', text)
    text = re.sub(r'</tr>', '\n', text)
    text = re.sub(r'</t[dh]>', '\t', text)
    text = re.sub(r'<[^>]+>', '', text)
    text = html_mod.unescape(text)
    return re.sub(r'\n{3,}', '\n\n', text).strip()


def _filetime_now() -> int:
    """Current time as Windows FILETIME (100ns intervals since 1601-01-01)."""
    return int((time.time() + 11644473600) * 10_000_000)


def _guess_mime(filename: str) -> str:
    return {
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.xls': 'application/vnd.ms-excel',
        '.pdf': 'application/pdf',
        '.csv': 'text/csv',
        '.txt': 'text/plain',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.zip': 'application/zip',
        '.doc': 'application/msword',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }.get(Path(filename).suffix.lower(), 'application/octet-stream')
