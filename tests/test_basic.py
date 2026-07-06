"""Tests for msgforge — verify with extract-msg."""

import extract_msg

from msgforge import Message


def test_plain_text(tmp_path):
    path = tmp_path / "test_plain.msg"
    msg = Message(
        subject="Plain Text Test",
        text_body="Hello, this is plain text.",
        to=[("alice@example.com", "Alice")],
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.subject == "Plain Text Test"
    assert "alice@example.com" in m.to
    m.close()


def test_html_body(tmp_path):
    path = tmp_path / "test_html.msg"
    msg = Message(
        subject="HTML Test",
        html_body="<html><body><p>Hello with <b>bold</b></p></body></html>",
        to=[("bob@example.com", "Bob")],
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.subject == "HTML Test"
    m.close()


def test_multiple_recipients(tmp_path):
    path = tmp_path / "test_multi.msg"
    msg = Message(
        subject="Multi Recip",
        text_body="Test",
        to=[("a@example.com", "Alice"), ("b@example.com", "Bob")],
        cc=[("c@example.com", "Carol")],
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert "a@example.com" in m.to
    assert "b@example.com" in m.to
    assert "c@example.com" in m.cc
    m.close()


def test_string_only_recipient(tmp_path):
    path = tmp_path / "test_string_recip.msg"
    msg = Message(
        subject="String Recip",
        text_body="Test",
        to=["user@example.com"],
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert "user@example.com" in m.to
    m.close()


def test_attachment(tmp_path):
    path = tmp_path / "test_attach.msg"
    msg = Message(
        subject="Attachment Test",
        text_body="See attached.",
        to=["user@example.com"],
    )
    msg.attach_bytes("data.txt", b"file content here")
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert len(m.attachments) == 1
    assert m.attachments[0].data == b"file content here"
    m.close()


def test_html_with_table_and_attachment(tmp_path):
    path = tmp_path / "test_full.msg"
    html = """<html><body>
    <p>Summary with <b>bold</b></p>
    <table border="1">
    <tr><th>Region</th><th>Revenue</th></tr>
    <tr><td>EMEA</td><td>$1.2M</td></tr>
    </table>
    </body></html>"""

    msg = Message(
        subject="Report",
        html_body=html,
        to=[("a@example.com", "Alice"), ("b@example.com", "Bob")],
    )
    msg.attach_bytes("report.xlsx", b"fake xlsx data")
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.subject == "Report"
    assert len(m.attachments) == 1
    m.close()


def test_chaining(tmp_path):
    path = tmp_path / "test_chain.msg"
    msg = (Message(subject="Chained", to=["x@example.com"])
           .attach_bytes("a.txt", b"aaa")
           .attach_bytes("b.txt", b"bbb"))
    msg.text_body = "Two attachments"
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert len(m.attachments) == 2
    m.close()


def test_inline_image(tmp_path):
    path = tmp_path / "test_inline.msg"
    png_header = b'\x89PNG\r\n\x1a\n' + b'\x00' * 100  # fake PNG
    msg = Message(
        subject="Inline Image",
        html_body='<p>Logo: <img src="cid:logo"></p>',
        to=["user@example.com"],
    )
    msg.attach_bytes("logo.png", png_header, content_id="logo")
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.subject == "Inline Image"
    assert len(m.attachments) == 1
    assert m.attachments[0].data == png_header
    m.close()


def test_inline_and_regular_attachments(tmp_path):
    path = tmp_path / "test_mixed.msg"
    msg = Message(
        subject="Mixed Attachments",
        html_body='<p>See image: <img src="cid:chart"></p>',
        to=["user@example.com"],
    )
    msg.attach_bytes("chart.png", b"fake png", content_id="chart")
    msg.attach_bytes("report.xlsx", b"fake xlsx")
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert len(m.attachments) == 2
    m.close()


def test_importance_high(tmp_path):
    path = tmp_path / "test_high.msg"
    msg = Message(
        subject="Urgent",
        text_body="Important!",
        to=["user@example.com"],
        importance="high",
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.subject == "Urgent"
    assert m.importance.value == 2  # HIGH
    m.close()


def test_importance_low(tmp_path):
    path = tmp_path / "test_low.msg"
    msg = Message(
        subject="FYI",
        text_body="No rush.",
        to=["user@example.com"],
        importance="low",
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.importance.value == 0  # LOW
    m.close()


def test_inline_image_data_uri_embedding(tmp_path):
    """Verify that cid: references are replaced with base64 data URIs in the RTF body."""
    import base64
    path = tmp_path / "test_data_uri.msg"
    png_data = b'\x89PNG\r\n\x1a\n' + b'\x00' * 50
    msg = Message(
        subject="Data URI Test",
        html_body='<html><body><img src="cid:pic"></body></html>',
        to=["user@example.com"],
    )
    msg.attach_bytes("pic.png", png_data, content_id="pic")
    msg.save(path)

    m = extract_msg.Message(str(path))
    # The RTF body should contain the base64-encoded image, not a cid: reference
    rtf_body = m.rtfBody
    b64 = base64.b64encode(png_data).decode('ascii')
    assert b64 in rtf_body.decode('ascii', errors='replace')
    assert b'cid:pic' not in rtf_body
    m.close()


def test_as_bytes(tmp_path):
    msg = Message(
        subject="Bytes Test",
        text_body="Hello",
        to=["user@example.com"],
    )
    data = msg.as_bytes()
    assert isinstance(data, bytes)
    assert len(data) > 0

    # Verify the bytes are valid by saving and re-reading
    path = tmp_path / "from_bytes.msg"
    path.write_bytes(data)
    m = extract_msg.Message(str(path))
    assert m.subject == "Bytes Test"
    m.close()


# ─── Regression & feature tests (0.3.0) ─────────────────────────────────────

def test_large_attachment_difat(tmp_path):
    """Files over ~7MB need DIFAT sectors; must survive strict OLE validation."""
    import olefile
    path = tmp_path / "test_big.msg"
    payload = b"\xAB" * (8 * 1024 * 1024)
    msg = Message(subject="Big", text_body="x", to=["user@example.com"])
    msg.attach_bytes("big.bin", payload)
    msg.save(path)

    ole = olefile.OleFileIO(str(path))  # validates FAT/DIFAT integrity
    ole.close()
    m = extract_msg.Message(str(path))
    assert m.attachments[0].data == payload
    m.close()


def test_whitespace_between_tags():
    """A space separating two tags must survive RTF encapsulation."""
    from msgforge._builder import _encapsulate_html
    rtf = _encapsulate_html('<b>bold</b> <i>italic</i>').decode()
    assert r'{\*\htmltag </b>} {\*\htmltag <i>}' in rtf


def test_newlines_in_text_become_spaces():
    """RTF readers ignore raw newlines — they must be converted to spaces."""
    from msgforge._builder import _encapsulate_html
    rtf = _encapsulate_html('<p>line one\nline two</p>').decode()
    assert 'line one line two' in rtf


def test_sender(tmp_path):
    path = tmp_path / "test_sender.msg"
    msg = Message(
        subject="From Me",
        text_body="Hello",
        to=["user@example.com"],
        sender=("boss@corp.example", "The Boss"),
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert "boss@corp.example" in m.sender
    assert "The Boss" in m.sender
    m.close()


def test_sent_with_datetime(tmp_path):
    from datetime import datetime
    path = tmp_path / "test_sent.msg"
    msg = Message(
        subject="Sent Mail",
        text_body="Hello",
        to=["user@example.com"],
        sent=datetime(2026, 7, 1, 12, 0, 0),
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.date is not None
    assert m.date.year == 2026 and m.date.month == 7 and m.date.day == 1
    m.close()


def test_html_body_property(tmp_path):
    """PidTagHtml (0x1013) should carry the raw HTML for non-Outlook readers."""
    path = tmp_path / "test_htmlprop.msg"
    html = "<p>hello <b>world</b></p>"
    msg = Message(subject="H", html_body=html, to=["user@example.com"])
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.htmlBody == html.encode("utf-8")
    m.close()


def test_invalid_importance_raises():
    import pytest
    with pytest.raises(ValueError):
        Message(importance="urgent")


def test_attachment_without_extension(tmp_path):
    path = tmp_path / "test_noext.msg"
    msg = Message(subject="NX", text_body="x", to=["user@example.com"])
    msg.attach_bytes("Makefile", b"all: build")
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert m.attachments[0].data == b"all: build"
    m.close()


def test_deterministic_output():
    def build():
        msg = Message(subject="Same", html_body="<p>x</p>", to=["user@example.com"])
        msg.attach_bytes("a.txt", b"data")
        return msg.as_bytes()
    assert build() == build()


def test_style_stripped_from_text_fallback(tmp_path):
    path = tmp_path / "test_style.msg"
    msg = Message(
        subject="Styled",
        html_body="<html><head><style>p { color: red; }</style></head>"
                  "<body><p>Visible</p></body></html>",
        to=["user@example.com"],
    )
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert "Visible" in m.body
    assert "color" not in m.body
    m.close()


def test_empty_display_name_falls_back_to_email(tmp_path):
    path = tmp_path / "test_emptyname.msg"
    msg = Message(subject="E", text_body="x", to=[("user@example.com", "")])
    msg.save(path)

    m = extract_msg.Message(str(path))
    assert "user@example.com" in m.to
    m.close()


def test_attach_filename_override_mime(tmp_path):
    """MIME type should be guessed from the effective (overridden) filename."""
    src = tmp_path / "data.bin"
    src.write_bytes(b"\x89PNG fake")
    msg = Message(subject="M", text_body="x", to=["user@example.com"])
    msg.attach(src, filename="image.png")
    assert msg._attachments[0][2] == "image/png"
