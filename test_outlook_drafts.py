import io
import json
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import pandas as pd

import email_sender
from email_sender import (
    ReportFile,
    generate_native_outlook_package,
    render_signature_html,
    render_subject,
)


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

    def test_native_package_contains_safe_importer_manifest_and_pdf(self):
        reports = [
            ReportFile(
                filename="Example.pdf",
                display_name="Example",
                normalized_name="example",
                content=b"%PDF-native-test",
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
                    "subject": "Your Pegasus Park Campus Engagement Report: January - June 2026",
                    "body": "Normal body",
                    "cc": "copy@example.org",
                    "bcc": "",
                }
            ]
        )

        with patch.object(email_sender, "_logo_bytes", return_value=b"test-logo"):
            archive, log = generate_native_outlook_package(
                rows,
                reports,
                "January - June 2026",
                "lhp7@lhholdings.net",
            )

        with ZipFile(io.BytesIO(archive)) as package:
            names = package.namelist()
            manifest = json.loads(package.read("drafts.json"))
            script = package.read("Create Outlook Drafts.ps1").decode("utf-8")
            pdf_path = manifest["drafts"][0]["attachment"]

            self.assertIn("START HERE - Create Outlook Drafts.cmd", names)
            self.assertIn("README - How to Create Outlook Drafts.txt", names)
            self.assertIn("signature_logo.png", names)
            self.assertIn(pdf_path, names)
            self.assertNotIn(".eml", " ".join(names))
            self.assertEqual(package.read(pdf_path), b"%PDF-native-test")

        self.assertEqual(manifest["target_mailbox"], "lhp7@lhholdings.net")
        self.assertEqual(manifest["draft_count"], 1)
        self.assertEqual(manifest["drafts"][0]["to"], "jane@example.org")
        self.assertEqual(manifest["drafts"][0]["cc"], ["copy@example.org"])
        self.assertIn("cid:water-cooler-signature-logo", manifest["drafts"][0]["html_body"])
        self.assertIn('$draftsFolder.Items.Add($olMailItem)', script)
        self.assertIn('$message.SendUsingAccount = $account', script)
        self.assertNotIn("$message.Send()", script)
        self.assertNotIn(".Send()", script)
        self.assertEqual(log.iloc[0]["Status"], "Packaged for native Outlook import")

    def test_native_package_rejects_invalid_mailbox(self):
        with self.assertRaisesRegex(ValueError, "valid Outlook sending mailbox"):
            generate_native_outlook_package(
                pd.DataFrame(),
                [],
                "January - June 2026",
                "not-an-email",
            )


if __name__ == "__main__":
    unittest.main()
