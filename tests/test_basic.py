"""Tests for msgforge — verify with extract-msg."""

from msgforge import Message
import extract_msg


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
