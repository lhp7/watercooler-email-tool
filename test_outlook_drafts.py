import io
import unittest
from email import policy
from email.parser import BytesParser
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

import pandas as pd

import email_sender
from email_sender import ReportFile, build_eml_message, generate_eml_zip, render_signature_html, render_subject


class OutlookDraftTests(unittest.TestCase):
    def test_default_subject_uses_requested_colon_format(self):
        subject = render_subject("Jane", "Example", "January - June 2026")

        self.assertEqual(subject, "Your Pegasus Park Campus Engagement Report: January - June 2026")

    def test_signature_matches_compact_reference_dimensions(self):
        html = render_signature_html(logo_cid="logo@example")

        self.assertIn('width="44"', html)
        self.assertIn("width:44px", html)
        self.assertIn("font-size:15px", html)
        self.assertIn("font-size:13px", html)
        self.assertEqual(html.count("font-size:11px"), 4)

    def test_eml_is_unsent_mailbox_neutral_and_complete(self):
        with patch.object(email_sender, "_logo_bytes", return_value=b"mime-test-logo"):
            message = build_eml_message(
                recipient_email="jane@example.org",
                recipient_name="Jane Doe",
                subject="Campus engagement report",
                body="Hi Jane,\n\nPlease review the attached report.",
                attachment_filename="Example Report.pdf",
                attachment_content=b"%PDF-test",
            )

        parsed = BytesParser(policy=policy.default).parsebytes(message.as_bytes())
        self.assertEqual(parsed["X-Unsent"], "1")
        self.assertIsNone(parsed["From"])
        self.assertIsNone(parsed["Sender"])
        self.assertEqual(parsed["To"].addresses[0].addr_spec, "jane@example.org")
        self.assertEqual(parsed["Subject"], "Campus engagement report")

        attachments = list(parsed.iter_attachments())
        self.assertEqual(len(attachments), 1)
        self.assertEqual(attachments[0].get_filename(), "Example Report.pdf")
        self.assertEqual(attachments[0].get_payload(decode=True), b"%PDF-test")

        html_parts = [part for part in parsed.walk() if part.get_content_type() == "text/html"]
        self.assertEqual(len(html_parts), 1)
        self.assertIn('width="44"', html_parts[0].get_content())
        related_images = [part for part in parsed.walk() if part.get_content_maintype() == "image"]
        self.assertEqual(len(related_images), 1)
        self.assertEqual(related_images[0].get_content_disposition(), "inline")

    def test_zip_contains_parseable_draft_with_normal_headers(self):
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

        with patch.object(email_sender, "SIGNATURE_LOGO_PATH", Path("missing-logo.png")):
            archive, log = generate_eml_zip(rows, reports, "January - June 2026")

        with ZipFile(io.BytesIO(archive)) as drafts:
            names = drafts.namelist()
            self.assertEqual(len(names), 1)
            self.assertTrue(names[0].endswith(".eml"))
            parsed = BytesParser(policy=policy.default).parsebytes(drafts.read(names[0]))

        self.assertEqual(parsed["To"].addresses[0].display_name, "Jane Doe")
        self.assertEqual(parsed["Subject"], "Normal subject")
        self.assertIsNone(parsed["From"])
        self.assertEqual(log.iloc[0]["Status"], "Draft generated")


if __name__ == "__main__":
    unittest.main()
