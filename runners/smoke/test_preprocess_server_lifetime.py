"""Real HTTP servers must survive the exit of the preprocess Python process.

Each test invokes the actual setup function in a child interpreter with tracked
assets copied into a temporary task. It waits for that interpreter to exit, then
fetches the page and compares its bytes. No database or model calls are made.
The original asyncio launch fails these tests in the Linux benchmark image;
macOS child watchers can mask that failure, so Linux execution is required.
"""
import os
from pathlib import Path
import shutil
import signal
import socket
import subprocess
import sys
import tarfile
import tempfile
import time
import unittest
from urllib.request import ProxyHandler, build_opener

ROOT = Path(__file__).resolve().parents[2]
CHILD = r'''
import asyncio, importlib.util, subprocess, sys
from pathlib import Path
main, root, port, pidfile, mode = sys.argv[1:]
real_popen = subprocess.Popen

def capture(*args, **kwargs):
    proc = real_popen(*args, **kwargs)
    argv = args[0] if args else kwargs.get('args', [])
    if isinstance(argv, (list, tuple)) and 'http.server' in argv:
        Path(pidfile).write_text(str(proc.pid))
    return proc

subprocess.Popen = capture
spec = importlib.util.spec_from_file_location('actual_preprocess', main)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
if mode == 'moex':
    module.PORT = int(port)
    asyncio.run(module.setup_mock_server())
else:
    module.TASK_ROOT = root
    asyncio.run(module.setup_mock_server(int(port)))
'''


class PreprocessServerLifetime(unittest.TestCase):
    def _stop_server(self, pid, port):
        proc = subprocess.run(['ps', '-p', str(pid), '-o', 'command='],
                              capture_output=True, text=True)
        if 'http.server' not in proc.stdout or str(port) not in proc.stdout.split():
            return
        try:
            os.kill(pid, signal.SIGTERM)
        except ProcessLookupError:
            return
        for _ in range(50):
            # In Docker this test can become the orphan's parent (PID 1).
            try:
                if os.waitpid(pid, os.WNOHANG)[0] == pid:
                    return
            except ChildProcessError:
                pass
            state = subprocess.run(['ps', '-p', str(pid), '-o', 'stat='],
                                   capture_output=True, text=True).stdout.strip()
            if not state or state.startswith('Z'):
                return
            time.sleep(.1)
        self.fail(f'HTTP test server {pid} did not stop after SIGTERM')

    def _check(self, task, mode, page):
        source = ROOT / 'tasks/finalpool' / task
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / 'preprocess').mkdir()
            main = root / 'preprocess/main.py'
            shutil.copy2(source / 'preprocess/main.py', main)
            if mode == 'moex':
                shutil.copytree(source / 'files/mock_pages', root / 'files/mock_pages')
                expected = (source / 'files/mock_pages' / page).read_bytes()
            else:
                (root / 'files').mkdir()
                archive = source / 'files/mock_pages.tar.gz'
                shutil.copy2(archive, root / 'files/mock_pages.tar.gz')
                with tarfile.open(archive) as tar:
                    expected = tar.extractfile('mock_pages/' + page).read()
            with socket.socket() as sock:
                sock.bind(('127.0.0.1', 0))
                port = sock.getsockname()[1]
            pidfile = root / 'server.pid'
            try:
                child = subprocess.run(
                    [sys.executable, '-c', CHILD, str(main), str(root),
                     str(port), str(pidfile), mode],
                    capture_output=True, text=True, timeout=60,
                )
            finally:
                if pidfile.exists():
                    self.addCleanup(self._stop_server, int(pidfile.read_text()), port)
            self.assertEqual(child.returncode, 0, child.stdout + child.stderr)
            self.assertTrue(pidfile.exists(), child.stdout + child.stderr)
            url = f'http://127.0.0.1:{port}/{page}'
            with build_opener(ProxyHandler({})).open(url, timeout=3) as response:
                self.assertEqual(response.status, 200)
                self.assertEqual(response.read(), expected)

    def test_moex_http_survives_preprocess_exit(self):
        self._check('moex-earnings-calendar-alert', 'moex', 'index.html')

    def test_terminal_http_survives_preprocess_exit(self):
        self._check('terminal-insales-pw-pricing-excel-word-gcal', 'terminal', 'pricing.html')


if __name__ == '__main__':
    unittest.main()
