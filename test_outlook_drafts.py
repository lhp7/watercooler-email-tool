import io
import json
import sys
import types
import unittest
from contextlib import nullcontext
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

import pandas as pd


def install_fake_msg_library():
    """Provide a small test double when the deployment dependency is unavailable."""
    module = types.ModuleType("independentsoft.msg")

    class Attachment:
        def __init__(self, file_path):
            self.file_name = Path(file_path).name
            self.data = Path(file_path).read_bytes()

    class Recipient:
        pass

    class Message:
        def __init__(self):
            self.message_flags = []
            self.store_support_masks = []
            self.recipients = []
            self.attachments = []

        def to_bytes(self):
            payload = {
                "subject": self.subject,
                "message_class": self.message_class,
                "message_flags": self.message_flags,
                "store_support_masks": self.store_support_masks,
                "display_to": self.display_to,
                "recipients": [
                    {
                        "name": recipient.display_name,
                        "email": recipient.email_address,
                        "type": recipient.recipient_type,
                    }
                    for recipient in self.recipients
                ],
                "attachments": [
                    {
                        "filename": attachment.file_name,
                        "data": attachment.data.decode("latin-1"),
                    }
                    for attachment in self.attachments
                ],
            }
            return json.dumps(payload).encode("utf-8")

        def save(self, file_path):
            Path(file_path).write_bytes(self.to_bytes())

    class DisplayType:
        MAIL_USER = "MAIL_USER"

    class MessageFlag:
        UNSENT = "UNSENT"

    class ObjectType:
        MAIL_USER = "MAIL_USER"

    class RecipientType:
        TO = "TO"
        CC = "CC"
        BCC = "BCC"

    class StoreSupportMask:
        CREATE = "CREATE"

    for item in (
        Attachment,
        DisplayType,
        Message,
        MessageFlag,
        ObjectType,
        Recipient,
        RecipientType,
        StoreSupportMask,
    ):
        setattr(module, item.__name__, item)

    package = types.ModuleType("independentsoft")
    package.msg = module
    sys.modules["independentsoft"] = package
    sys.modules["independentsoft.msg"] = module


install_fake_msg_library()

import email_sender
from email_sender import ReportFile, build_msg_message, generate_msg_zip, render_signature_html
from independentsoft.msg import MessageFlag, RecipientType, StoreSupportMask


class OutlookDraftTests(unittest.TestCase):
    def test_signature_logo_is_small_in_attribute_and_inline_style(self):
        with patch.object(email_sender, "_signature_logo_data_uri", return_value="data:image/png;base64,dGVzdA=="):
            html = render_signature_html()

        self.assertIn('width="56"', html)
        self.assertIn("width:56px", html)

    def test_msg_is_unsent_mailbox_neutral_and_complete(self):
        attachment_path = Path(__file__).with_name("_test_example_report.pdf")
        attachment_path.write_bytes(b"%PDF-test")
        try:
            message = build_msg_message(
                recipient_email="jane@example.org",
                recipient_name="Jane Doe",
                subject="Campus engagement report",
                text_body="Plain body",
                html_body="<p>HTML body</p>",
                attachment_filename="Example Report.pdf",
                attachment_file_path=attachment_path,
                cc="copy@example.org",
            )
        finally:
            attachment_path.unlink(missing_ok=True)

        self.assertEqual(message.message_class, "IPM.Note")
        self.assertIn(MessageFlag.UNSENT, message.message_flags)
        self.assertIn(StoreSupportMask.CREATE, message.store_support_masks)
        self.assertEqual(message.subject, "Campus engagement report")
        self.assertEqual(message.recipients[0].email_address, "jane@example.org")
        self.assertEqual(message.recipients[0].recipient_type, RecipientType.TO)
        self.assertEqual(message.recipients[1].recipient_type, RecipientType.CC)
        self.assertEqual(message.attachments[0].file_name, "Example Report.pdf")
        self.assertEqual(message.attachments[0].data, b"%PDF-test")
        self.assertIn(b"fromhtml1", message.body_rtf)
        self.assertFalse(hasattr(message, "sender_email_address"))
        self.assertFalse(hasattr(message, "sender_name"))

    def test_zip_contains_native_msg_draft_with_normal_fields(self):
        reports = [
            ReportFile(
                filename="Example.pdf",
                display_name="Example",
                normalized_name="example",
                content=b"%PDF-test",
            )
        ]
        rows = pd.DataFrame(
            [
                {
                    "include": True,
                    "status": "Ready",
                    "first_name": "Jane",
                    "last_name": "Doe",
                    "recipient_name": "Jane Doe",
                    "email": "jane@example.org",
                    "org_name": "Example",
                    "matched_report": "Example.pdf",
                    "match_score": 100,
                    "subject": "Normal subject",
                    "body": "Normal body",
                    "cc": "",
                    "bcc": "",
                }
            ]
        )

        work_dir = Path(__file__).parent
        generated_attachment = work_dir / "Example Campus Engagement Report January - June 2026.pdf"
        generated_draft = work_dir / "example_jane_doe.msg"
        try:
            with (
                patch.object(email_sender, "SIGNATURE_LOGO_PATH", Path("missing-logo.png")),
                patch.object(email_sender, "TemporaryDirectory", return_value=nullcontext(str(work_dir))),
            ):
                archive, log = generate_msg_zip(rows, reports, "January - June 2026")
        finally:
            generated_attachment.unlink(missing_ok=True)
            generated_draft.unlink(missing_ok=True)

        with ZipFile(io.BytesIO(archive)) as drafts:
            names = drafts.namelist()
            self.assertEqual(len(names), 1)
            self.assertTrue(names[0].endswith(".msg"))
            self.assertNotIn(".eml", names[0])
            payload = json.loads(drafts.read(names[0]))

        self.assertEqual(payload["subject"], "Normal subject")
        self.assertEqual(payload["recipients"][0]["name"], "Jane Doe")
        self.assertEqual(payload["recipients"][0]["email"], "jane@example.org")
        self.assertIn("UNSENT", payload["message_flags"])
        self.assertIn("CREATE", payload["store_support_masks"])
        self.assertEqual(payload["attachments"][0]["filename"], "Example Campus Engagement Report January - June 2026.pdf")
        self.assertEqual(log.iloc[0]["Status"], "Draft generated")


if __name__ == "__main__":
    unittest.main()
