"""
Preprocess script for sf-ticket-escalation-report task.

1. Clears email data and prior Teamly escalation pages (writable schemas).
2. Ensures a Teamly space exists for the agent to create the page under.
3. Extracts mock_pages.tar.gz and starts HTTP server on port 30210.
4. sf_data is read-only -- do NOT modify (realia russified centrally by
   db/zzz_clickhouse_after_init.sql).
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
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_emails(cur):
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] Email data cleared.")


def setup_teamly(cur):
    """Ensure a Teamly space exists for the agent to create the escalation page
    under, and clear any prior escalation pages (idempotency).

    We intentionally do NOT pre-create the answer page -- the agent must produce
    it so the evaluation actually exercises the agent's work.
    """
    print("[preprocess] Setting up Teamly...")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; "
              "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return
    # Dedicated space for the agent to drop the escalation analysis page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('SUPPORT', 'Команда поддержки',
                'База знаний и аналитические отчёты команды поддержки.')
        ON CONFLICT (key) DO NOTHING
    """)
    # Idempotency: drop any escalation pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%escalation%'
            OR title ILIKE '%эскалац%'
    """)
    print("[preprocess] Teamly ready: 'SUPPORT' space ensured, prior pages cleared.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30210."""
    print("[preprocess] Setting up mock escalation rules API...")

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
    print(f"[preprocess] Mock API server running at http://localhost:{port}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_emails(cur)
        setup_teamly(cur)
        conn.commit()
        print("[preprocess] DB operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()
    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
