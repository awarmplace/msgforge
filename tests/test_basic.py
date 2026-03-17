"""Tests for msgforge — verify with extract-msg."""

from msgforge import Message
import extract_msg


def test_plain_text():
    msg = Message(
        subject="Plain Text Test",
        text_body="Hello, this is plain text.",
        to=[("alice@example.com", "Alice")],
    )
    msg.save("test_plain.msg")

    m = extract_msg.Message("test_plain.msg")
    assert m.subject == "Plain Text Test"
    assert "alice@example.com" in m.to
    m.close()


def test_html_body():
    msg = Message(
        subject="HTML Test",
        html_body="<html><body><p>Hello with <b>bold</b></p></body></html>",
        to=[("bob@example.com", "Bob")],
    )
    msg.save("test_html.msg")

    m = extract_msg.Message("test_html.msg")
    assert m.subject == "HTML Test"
    m.close()


def test_multiple_recipients():
    msg = Message(
        subject="Multi Recip",
        text_body="Test",
        to=[("a@example.com", "Alice"), ("b@example.com", "Bob")],
        cc=[("c@example.com", "Carol")],
    )
    msg.save("test_multi.msg")

    m = extract_msg.Message("test_multi.msg")
    assert "a@example.com" in m.to
    assert "b@example.com" in m.to
    assert "c@example.com" in m.cc
    m.close()


def test_string_only_recipient():
    msg = Message(
        subject="String Recip",
        text_body="Test",
        to=["user@example.com"],
    )
    msg.save("test_string_recip.msg")

    m = extract_msg.Message("test_string_recip.msg")
    assert "user@example.com" in m.to
    m.close()


def test_attachment():
    msg = Message(
        subject="Attachment Test",
        text_body="See attached.",
        to=["user@example.com"],
    )
    msg.attach_bytes("data.txt", b"file content here")
    msg.save("test_attach.msg")

    m = extract_msg.Message("test_attach.msg")
    assert len(m.attachments) == 1
    assert m.attachments[0].data == b"file content here"
    m.close()


def test_html_with_table_and_attachment():
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
    msg.save("test_full.msg")

    m = extract_msg.Message("test_full.msg")
    assert m.subject == "Report"
    assert len(m.attachments) == 1
    m.close()


def test_chaining():
    msg = (Message(subject="Chained", to=["x@example.com"])
           .attach_bytes("a.txt", b"aaa")
           .attach_bytes("b.txt", b"bbb"))
    msg.text_body = "Two attachments"
    msg.save("test_chain.msg")

    m = extract_msg.Message("test_chain.msg")
    assert len(m.attachments) == 2
    m.close()
