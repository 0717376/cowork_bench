"""Read real temporary files through PG backend, with an in-memory DB receipt."""
import os
import sys
import tempfile
import unittest
from email import message_from_string
from pathlib import Path
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from emails_mcp.server import create_services, config_manager
from emails_mcp.models.config import EmailConfig, WorkspaceConfig
from emails_mcp.utils.exceptions import SendEmailError


class Receipt:
    def __init__(self):
        self.statements = []
    def cursor(self):
        return self
    def __enter__(self):
        return self
    def __exit__(self, *args):
        pass
    def execute(self, query, params=None):
        self.statements.append((query, params))
    def fetchone(self):
        return (123,)


class AttachmentRootTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name) / "workspace"
        self.root.mkdir()
        self.other = Path(self.tmp.name) / "server-cwd"
        self.other.mkdir()
        self.previous_cwd = Path.cwd()
        self.addCleanup(os.chdir, self.previous_cwd)
        os.chdir(self.other)
        self.data = b"synthetic attachment bytes: \x00\xff"
        (self.root / "Report.xlsx").write_bytes(self.data)
        # A same-name cwd file proves there is no fallback to the server directory.
        (self.other / "Report.xlsx").write_bytes(b"wrong cwd attachment")
        self.guard = patch("psycopg2.connect", side_effect=AssertionError("real DB forbidden"))
        self.guard.start()
        self.addCleanup(self.guard.stop)
        with patch.object(config_manager, "workspace_config", WorkspaceConfig(
                attachment_upload_path=str(self.root))):
            self.service, *_ = create_services(EmailConfig("sender@example.invalid", "unused"))
        self.backend = self.service.smtp_backend
        self.receipt = Receipt()
        self.backend.connection = self.receipt
        self.backend.ensure_connected = Mock()

    def send(self, paths):
        return self.backend.send_email("recipient@example.invalid", "synthetic", "body", attachments=paths)

    def assert_attachment(self, path, filename="Report.xlsx"):
        sent, raw = self.send([path])
        self.assertTrue(sent)
        parts = list(message_from_string(raw).walk())
        attachments = [p for p in parts if p.get_content_disposition() == "attachment"]
        self.assertEqual(len(attachments), 1)
        self.assertEqual(attachments[0].get_payload(decode=True), self.data)
        self.assertEqual(attachments[0].get_filename(), filename)
        self.assertEqual(attachments[0].get_content_type(), "application/octet-stream")
        rows = [p for q, p in self.receipt.statements if "INSERT INTO email.attachments" in q]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][1:4], (filename, "application/octet-stream", len(self.data)))
        self.assertIn("INSERT INTO email.sent_log", self.receipt.statements[-1][0])

    def assert_rejected(self, path):
        with self.assertRaises(SendEmailError):
            self.send(["Report.xlsx", path])
        self.assertEqual(self.receipt.statements, [])
        self.backend.ensure_connected.assert_not_called()

    def test_relative_path_uses_configured_root_not_cwd(self):
        self.assert_attachment("Report.xlsx")

    def test_absolute_path_inside_root_is_supported(self):
        self.assert_attachment(str(self.root / "Report.xlsx"))

    def test_nested_relative_path(self):
        (self.root / "nested").mkdir()
        (self.root / "nested" / "Report.xlsx").write_bytes(self.data)
        self.assert_attachment("nested/Report.xlsx")

    def test_missing_file_rejects_before_any_database_work(self):
        self.assert_rejected("Missing.xlsx")

    def test_directory_is_not_attachment(self):
        self.assert_rejected(str(self.root))

    def test_parent_escape_is_rejected(self):
        self.assert_rejected("../server-cwd/Report.xlsx")

    def test_absolute_outside_root_is_rejected(self):
        self.assert_rejected(str(self.other / "Report.xlsx"))

    def test_symlink_escape_is_rejected(self):
        (self.root / "escape.xlsx").symlink_to(self.other / "Report.xlsx")
        self.assert_rejected("escape.xlsx")

    def test_internal_symlink_keeps_requested_filename_and_content(self):
        (self.root / "Alias.xlsx").symlink_to(self.root / "Report.xlsx")
        self.assert_attachment("Alias.xlsx", "Alias.xlsx")

    def test_no_attachment_send_keeps_existing_behavior(self):
        sent, raw = self.send(None)
        self.assertTrue(sent)
        self.assertEqual(message_from_string(raw).get_payload(decode=True), b"body")
        self.assertEqual(len(self.receipt.statements), 3)

    def test_unconfigured_standalone_backend_keeps_cwd_semantics(self):
        from emails_mcp.backends.pg_backend import PgSMTPBackend
        standalone = PgSMTPBackend(EmailConfig("sender@example.invalid", "unused"))
        standalone.connection = Receipt()
        standalone.ensure_connected = Mock()
        sent, raw = standalone.send_email("recipient@example.invalid", "synthetic", "body",
                                         attachments=["Report.xlsx"])
        self.assertTrue(sent)
        attachment = next(p for p in message_from_string(raw).walk()
                          if p.get_content_disposition() == "attachment")
        self.assertEqual(attachment.get_payload(decode=True), b"wrong cwd attachment")

    def test_service_uses_same_backend_and_stores_success(self):
        self.assertTrue(self.service.send_email("recipient@example.invalid", "synthetic", "body",
                                              attachments=["Report.xlsx"]))
        self.assertTrue(any("email.attachments" in q for q, _ in self.receipt.statements))


if __name__ == "__main__":
    unittest.main()
