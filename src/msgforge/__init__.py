"""msgforge — Pure Python Outlook .msg file builder.

Create Outlook .msg files with HTML bodies, recipients, and attachments.
No COM, no Outlook installation, no native dependencies. Works on Linux.

Usage::

    from msgforge import Message

    msg = Message(
        subject="Q1 Report",
        html_body="<p>Please see <b>attached</b> report.</p>",
        to=[("alice@example.com", "Alice")],
    )
    msg.attach("report.xlsx")
    msg.save("Q1 Report.msg")
"""

from ._builder import Message

__all__ = ["Message"]
__version__ = "0.1.1"
