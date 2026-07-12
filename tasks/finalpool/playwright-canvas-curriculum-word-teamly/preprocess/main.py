"""
Preprocess script for playwright-canvas-curriculum-word task (teamly variant).

This script:
1. Ensures a Teamly space exists, clears leftover tracker pages, clears email data
2. Injects noise data (a noise Teamly page + a noise email)
3. Extracts mock_pages.tar.gz and starts HTTP server on port 30210
NOTE: Does NOT modify Canvas data (read-only).
NOTE: Does NOT pre-create the 'Course Compliance Tracker' page or its 22 rows —
      the agent must produce those itself.
"""

import argparse
import asyncio
import json
import os
import shutil
import tarfile
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


SPACE_KEY = "ACCRED"
SPACE_NAME = "Аккредитация"
SPACE_DESC = "Рабочее пространство комиссии по аккредитации учебных программ."


def setup_teamly_space(cur):
    """Ensure a Teamly space exists and clear leftover tracker pages (idempotency).

    We intentionally do NOT pre-create the 'Course Compliance Tracker' page —
    the agent must create it so the evaluation actually tests the agent.
    """
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; "
              "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return None
    cur.execute(
        """INSERT INTO teamly.spaces (key, name, description)
           VALUES (%s, %s, %s)
           ON CONFLICT (key) DO NOTHING""",
        (SPACE_KEY, SPACE_NAME, SPACE_DESC),
    )
    # Idempotency: drop any tracker pages left from previous runs.
    cur.execute(
        """DELETE FROM teamly.pages
            WHERE title ILIKE '%%course compliance tracker%%'
               OR title ILIKE '%%compliance tracker%%'
               OR (title ILIKE '%%course%%' AND title ILIKE '%%compliance%%')"""
    )
    cur.execute("SELECT id FROM teamly.spaces WHERE key = %s", (SPACE_KEY,))
    row = cur.fetchone()
    print("[preprocess] Teamly ready: space ensured, prior tracker pages cleared.")
    return row[0] if row else None


def clear_emails(cur):
    """Clear all email data."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    try:
        cur.execute("DELETE FROM email.sent_log")
    except Exception:
        pass
    cur.execute("DELETE FROM email.messages")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] Email data cleared.")


def inject_noise_teamly(cur, space_id):
    """Inject a noise Teamly page (unrelated to the tracker)."""
    if space_id is None:
        return
    print("[preprocess] Injecting noise Teamly page...")
    cur.execute(
        "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
        (
            space_id,
            "Протокол заседания — февраль",
            "Краткие заметки с февральского заседания кафедры. "
            "Обсудили расписание и хозяйственные вопросы. К аккредитации не относится.",
            "Секретариат",
        ),
    )
    print("[preprocess] Noise Teamly page injected.")


def inject_noise_email(cur):
    """Inject noise email data."""
    print("[preprocess] Injecting noise email data...")
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if not row:
        return
    folder_id = row[0]
    cur.execute(
        """INSERT INTO email.messages
            (folder_id, message_id, subject, from_addr, to_addr, body_text)
           VALUES (%s, %s, %s, %s, %s, %s)""",
        (
            folder_id,
            f"<noise-{uuid.uuid4()}@university.edu>",
            "Обновление по хозяйственной части кампуса",
            "facilities@university.edu",
            json.dumps(["all-staff@university.edu"]),
            "Ремонт библиотеки будет завершён к 15 марта.",
        ),
    )
    print("[preprocess] Noise email data injected.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30210."""
    print("[preprocess] Setting up mock accreditation standards server...")

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    port = 30210

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --directory \"{mock_dir}\" "
        f"> \"{mock_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{port}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        space_id = setup_teamly_space(cur)
        clear_emails(cur)
        inject_noise_teamly(cur, space_id)
        inject_noise_email(cur)
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
