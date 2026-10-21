"""Preprocess for yf-earnings-calendar-alert.
Clears email schema and starts mock web server on port 30225."""
import argparse
import asyncio
import os
import shutil
import subprocess
import sys

import psycopg2

PORT = 30225
DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

# Handle of the detached http.server; kept referenced so the Popen object is
# not finalized while this process lives (it only warns, but stay explicit).
_server_proc = None


def clear_email():
    conn = psycopg2.connect(**DB)
    conn.autocommit = True
    cur = conn.cursor()
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    # Ensure sent folder exists
    cur.execute("SELECT id FROM email.folders WHERE LOWER(name) LIKE '%sent%' LIMIT 1")
    if not cur.fetchone():
        cur.execute("INSERT INTO email.folders (name) VALUES ('Sent')")
    cur.close()
    conn.close()
    print("[preprocess] Cleared email tables.")


async def run_command(cmd):
    proc = await asyncio.create_subprocess_shell(
        cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    await proc.wait()


async def setup_mock_server():
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    mock_dir = os.path.join(task_root, "files", "mock_pages")
    tmp_dir = os.path.join(task_root, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    await run_command(f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null")
    await asyncio.sleep(0.5)
    log_path = os.path.join(tmp_dir, "server.log")
    # The http.server must outlive this preprocess process. An asyncio
    # subprocess transport owns the child and SIGKILLs it when the
    # asyncio.run loop is torn down (start_new_session does not prevent
    # that), so launch it fully detached via subprocess.Popen in its own
    # session instead.
    global _server_proc
    with open(log_path, "wb") as log_fh:
        _server_proc = subprocess.Popen(
            [sys.executable or "python3", "-m", "http.server", str(PORT),
             "--directory", mock_dir],
            stdin=subprocess.DEVNULL, stdout=log_fh, stderr=log_fh,
            start_new_session=True,
        )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{PORT} (pid={_server_proc.pid})")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    try:
        clear_email()
    except Exception as e:
        print(f"[preprocess] clear_email warning: {e}")
    try:
        await setup_mock_server()
    except Exception as e:
        print(f"[preprocess] setup_mock_server warning: {e}")
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
