"""
Preprocess script for playwright-clickhouse-competitor-analysis-teamly-excel task.

1. Clears Teamly user pages and email data (writable schemas).
2. Injects noise data into email and Teamly.
3. Extracts mock_pages.tar.gz and starts HTTP server on port 30216.
4. sf_data (ClickHouse-identity warehouse) is read-only -- do NOT modify; its
   region/segment data values are russified centrally by
   db/zzz_clickhouse_after_init.sql.
"""
import argparse
import asyncio
import glob as globmod
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


def clear_teamly(cur):
    """Clear user-created Teamly pages and ensure a target space exists.

    Global seed pages have id <= 3; only user-created pages are removed.
    No competitor page/space deliverable is pre-created -- the agent must
    produce the 'Competitive Intelligence Tracker' content itself.
    """
    print("[preprocess] Clearing Teamly data...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute(
            """
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('STRATEGY', 'Отдел стратегии',
                    'Пространство для аналитики конкурентного позиционирования.')
            ON CONFLICT (key) DO NOTHING
            """
        )
    print("[preprocess] Teamly data cleared.")


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


def inject_noise_emails(cur):
    """Insert noise emails so the agent must filter correctly."""
    print("[preprocess] Injecting noise email data...")
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if not row:
        return
    folder_id = row[0]
    noise = [
        ("Заметки еженедельной планёрки команды", "pm@company.com", '["team@company.com"]',
         "Итоги планёрки за неделю. Блокеров не зафиксировано."),
        ("Заказ канцелярских товаров", "admin@company.com", '["procurement@company.com"]',
         "Нужно дозаказать бумагу для принтера и картриджи с тонером."),
        ("Напоминание о квартальной аттестации", "hr@company.com", '["all-staff@company.com"]',
         "Напоминаем: оценки эффективности нужно сдать до конца месяца."),
    ]
    for subj, from_addr, to_addr, body in noise:
        cur.execute("""
            INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, body_text)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (folder_id, f"<noise-{uuid.uuid4()}@company.com>", subj, from_addr, to_addr, body))
    print("[preprocess] Noise emails injected.")


def inject_noise_teamly(cur):
    """Inject a noise Teamly page (distinct from the deliverable)."""
    print("[preprocess] Injecting noise Teamly data...")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        return
    cur.execute("SELECT id FROM teamly.spaces WHERE key='STRATEGY' LIMIT 1")
    row = cur.fetchone()
    if not row:
        return
    space_id = row[0]
    cur.execute("""
        INSERT INTO teamly.pages (space_id, title, body, author)
        VALUES (%s, %s, %s, %s)
    """, (
        space_id,
        "Трекер проектов",
        "# Трекер проектов\n\nСпринт Q1: задачи команды разработки. К конкурентному анализу не относится.",
        "admin",
    ))
    print("[preprocess] Noise Teamly data injected.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30216."""
    print("[preprocess] Setting up mock competitor intelligence server...")

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
    port = 30216

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
        clear_teamly(cur)
        clear_emails(cur)
        inject_noise_emails(cur)
        inject_noise_teamly(cur)
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
        for pattern in ["Competitive_Analysis.xlsx"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
