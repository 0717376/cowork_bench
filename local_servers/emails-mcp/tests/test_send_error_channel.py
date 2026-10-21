"""Exercise registered send_email over the real MCP protocol, with no SMTP/DB."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from mcp.server.fastmcp import FastMCP
from mcp.shared.memory import create_connected_server_and_client_session
from emails_mcp.tools.email_tools import register_email_tools
from emails_mcp.services.email_service import EmailService
from emails_mcp.models.config import EmailConfig
from emails_mcp.utils.exceptions import SendEmailError


class SendChannelTests(unittest.IsolatedAsyncioTestCase):
    async def call(self, outcome):
        class Service:
            def send_email(self, **kwargs):
                if isinstance(outcome, Exception):
                    raise outcome
                return outcome
        server = FastMCP("email-contract-test")
        register_email_tools(server, Service())
        async with create_connected_server_and_client_session(server._mcp_server) as client:
            result = await client.call_tool("send_email", {
                "to": "recipient@example.invalid", "subject": "synthetic", "body": "body"
            })
            # The client parses the wire result, rather than inspecting a handler string.
            return result.model_dump(mode="json")

    async def test_exception_is_protocol_error_with_cause(self):
        result = await self.call(FileNotFoundError("synthetic missing attachment"))
        self.assertTrue(result["isError"])
        self.assertIn("synthetic missing attachment", str(result["content"]))

    async def test_false_is_protocol_error(self):
        result = await self.call(False)
        self.assertTrue(result["isError"])
        self.assertIn("Email sending failed", str(result["content"]))

    async def test_success_remains_success(self):
        result = await self.call(True)
        self.assertFalse(result["isError"])
        self.assertIn("Email sent successfully", str(result["content"]))

    async def test_real_service_preserves_backend_failure_channel(self):
        service = EmailService(EmailConfig("sender@example.invalid", "unused"))
        service.smtp_backend.send_email = lambda **kw: (_ for _ in ()).throw(
            SendEmailError("synthetic rejected attachment"))
        server = FastMCP("real-service-contract-test")
        register_email_tools(server, service)
        with patch("psycopg2.connect", side_effect=AssertionError("real DB forbidden")):
            async with create_connected_server_and_client_session(server._mcp_server) as client:
                result = await client.call_tool("send_email", {
                    "to": "recipient@example.invalid", "subject": "synthetic", "body": "body"
                })
        self.assertTrue(result.isError)
        self.assertIn("synthetic rejected attachment", str(result.content))


if __name__ == "__main__":
    unittest.main()
