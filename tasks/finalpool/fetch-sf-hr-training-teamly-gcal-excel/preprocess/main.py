"""
Preprocess script for fetch-sf-hr-training-clickhouse-gcal-excel task.

1. Clears teamly user pages, gcal, email data (writable schemas).
2. Injects noise data.
3. Extracts mock_pages.tar.gz and starts HTTP server on port 30219.
4. sf_data (ClickHouse warehouse) is read-only -- do NOT modify.
"""
import argparse
import asyncio
import glob as globmod
import json
import os
import shutil
import tarfile
import uuid
from datetime import datetime, timedelta

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

TRAINING_SPACE_KEY = "TRAINING"


def clear_teamly(cur):
    """Drop user-created Teamly pages (seed pages have id <= 3); ensure a space
    exists for the agent to use. Does NOT pre-seed the training program pages."""
    print("[preprocess] Clearing Teamly user data...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES (%s, %s, %s)
            ON CONFLICT (key) DO NOTHING
        """, (
            TRAINING_SPACE_KEY,
            "Обучение и развитие",
            "База знаний по программам обучения и развития персонала.",
        ))
    print("[preprocess] Teamly user data cleared.")


def clear_gcal(cur):
    print("[preprocess] Clearing GCal events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] GCal events cleared.")


def clear_emails(cur):
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


def inject_noise_teamly(cur):
    """Inject a noise Teamly page the agent must ignore (must NOT satisfy the
    training-program check)."""
    print("[preprocess] Injecting noise Teamly data...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is None:
        return
    cur.execute("SELECT id FROM teamly.spaces WHERE key = %s", (TRAINING_SPACE_KEY,))
    row = cur.fetchone()
    if row is None:
        cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
        row = cur.fetchone()
    if row is None:
        return
    space_id = row[0]
    cur.execute(
        "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
        (space_id, "Справочник сотрудников",
         "Старый справочник контактов отдела кадров. Не относится к текущей задаче.",
         "hr"),
    )
    print("[preprocess] Noise Teamly data injected.")


def inject_noise_emails(cur):
    """Insert noise emails."""
    print("[preprocess] Injecting noise email data...")
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if not row:
        return
    folder_id = row[0]
    noise = [
        ("Напоминание об открытой регистрации льгот", "benefits@company.com", '["all-staff@company.com"]',
         "Регистрация на льготы 2026 года закрывается 31 марта."),
        ("Требуется тренинг по информационной безопасности", "it-security@company.com", '["all-staff@company.com"]',
         "Ежегодный тренинг по информационной безопасности теперь доступен."),
        ("Уведомление о ремонте парковки", "facilities@company.com", '["all-staff@company.com"]',
         "Стоянка B будет закрыта на перекладку покрытия с 1 по 3 апреля."),
    ]
    for subj, from_addr, to_addr, body in noise:
        cur.execute("""
            INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, body_text)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (folder_id, f"<noise-{uuid.uuid4()}@company.com>", subj, from_addr, to_addr, body))
    print("[preprocess] Noise emails injected.")


def inject_noise_gcal(cur, launch_dt):
    """Inject noise calendar events."""
    print("[preprocess] Injecting noise GCal events...")
    dt = (launch_dt + timedelta(days=55)).strftime('%Y-%m-%d')
    cur.execute(f"""
        INSERT INTO gcal.events (summary, description, start_datetime, end_datetime)
        VALUES ('Общее собрание', 'Ежеквартальное общее собрание', '{dt} 10:00:00+00', '{dt} 11:00:00+00')
    """)
    print("[preprocess] Noise GCal events injected.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30219."""
    print("[preprocess] Setting up mock training catalog server...")

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
    port = 30219

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

    launch_dt = datetime(2026, 3, 7)
    if args.launch_time:
        try:
            # tolerate weekday tail, e.g. "2026-03-07 10:00:00 Saturday"
            launch_dt = datetime.strptime(" ".join(args.launch_time.split()[:2]), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass

    try:
        clear_teamly(cur)
        clear_gcal(cur)
        clear_emails(cur)
        inject_noise_teamly(cur)
        inject_noise_emails(cur)
        inject_noise_gcal(cur, launch_dt)
        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()

    if args.agent_workspace:
        for pattern in ["Training_Budget_Plan.xlsx"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
