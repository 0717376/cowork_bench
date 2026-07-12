"""
Preprocess script for fetch-notion-monitoring task.

This script:
1. Clears Teamly (user pages), Google Calendar, and email data; ensures an empty
   MONITORING space exists (does NOT pre-create the target dashboard page)
2. Extracts mock_pages.tar.gz and starts HTTP server on port 30154
"""

import argparse
import asyncio
import os
import shutil
import tarfile

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_teamly(cur):
    """Clear user-created Teamly pages and ensure an empty MONITORING space.

    Seed pages have id <= 3; only user-created pages (id > 3) are removed so the
    workspace starts clean. We do NOT pre-create the target dashboard page.
    """
    print("[preprocess] Clearing Teamly data and ensuring MONITORING space...")
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('MONITORING', 'Мониторинг инфраструктуры',
                        'База знаний команды инфраструктуры: отчёты о доступности сервисов.')
                ON CONFLICT (key) DO NOTHING
                """
            )
    except Exception as e:
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")
    print("[preprocess] Teamly data cleared.")


def clear_gcal(cur):
    """Clear all Google Calendar events."""
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] Google Calendar events cleared.")


def clear_emails(cur):
    """Clear all email data."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] Email data cleared.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30154."""
    print("[preprocess] Setting up mock status server...")

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    port = 30154

    # Kill any existing process on the port
    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    # Start HTTP server
    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --directory \"{mock_dir}\" "
        f"> \"{mock_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock status server running at http://localhost:{port}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_teamly(cur)
        clear_gcal(cur)
        clear_emails(cur)
        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
