"""Real terminal execution: quoted punctuation is data; operators stay gated."""

import importlib.util
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "local_servers/cli-mcp-server/src/cli_mcp_server/server.py"
spec = importlib.util.spec_from_file_location("cowork_terminal", SRC)
terminal = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = terminal
with patch.dict(os.environ, {"ALLOWED_DIR": str(ROOT)}):
    spec.loader.exec_module(terminal)


class TerminalQuotedArguments(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.workspace = Path(self.temp.name) / "workspace"
        self.workspace.mkdir()
        self.config = terminal.SecurityConfig(
            {"python", "python3", "echo", "cat"},
            set(),
            2048,
            5,
            allow_all_flags=True,
            allow_shell_operators=True,
        )
        self.executor = terminal.CommandExecutor(str(self.workspace), self.config)

    def test_python_semicolon_comparison_and_pipe_in_quotes_are_one_argument(self):
        command = '''python3 -c "import json; print(json.dumps({'text':'a|b;c', 'ok': 2 > 1}))"'''
        result = self.executor.execute(command)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("a|b;c", result.stdout)
        self.assertIn("true", result.stdout)

    def test_literal_operators_need_no_shell_permission(self):
        self.config.allow_shell_operators = False
        self.assertEqual(
            self.executor.execute("echo 'a;b|c>d'").stdout.strip(), "a;b|c>d"
        )
        self.assertEqual(self.executor.execute(r"echo a\;b").stdout.strip(), "a;b")
        with self.assertRaises(terminal.CommandSecurityError):
            self.executor.execute("echo a; echo b")

    def test_real_pipeline_and_redirect_still_validate_commands_and_paths(self):
        self.assertEqual(
            self.executor.execute(
                "echo 'a;b' > 'file with spaces.txt' && cat 'file with spaces.txt'"
            ).stdout.strip(),
            "a;b",
        )
        with self.assertRaises(terminal.CommandSecurityError):
            self.executor.execute("echo a; forbidden-command")
        sibling = Path(self.temp.name) / "workspace-other"
        sibling.mkdir()
        for cmd in [f'echo a > "{sibling}/escaped.txt"', f'cat "{sibling}/data.txt"']:
            with (
                self.subTest(cmd=cmd),
                self.assertRaises(terminal.CommandSecurityError),
            ):
                self.executor.execute(cmd)
        self.assertFalse((sibling / "escaped.txt").exists())

    def test_nested_execution_is_not_an_allowlist_escape(self):
        for command in [
            'echo first && echo "$(forbidden-command)"',
            "echo first && echo `forbidden-command`",
            "echo first && echo a & forbidden-command",
            "echo first && echo a\nforbidden-command",
        ]:
            with (
                self.subTest(command=command),
                self.assertRaises(terminal.CommandSecurityError),
            ):
                self.executor.execute(command)
        self.assertEqual(
            self.executor.execute("echo '$(literal);not-executed'").stdout.strip(),
            "$(literal);not-executed",
        )


if __name__ == "__main__":
    unittest.main()
