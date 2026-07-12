"""Preprocess script for sf-teamly-project-tracker-excel-gcal-email."""
import os
import argparse, json, os, sys, shutil, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Noise page must NOT satisfy the dashboard check.
NOISE_PAGE_TITLE = "Архив протоколов совещаний"

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM gcal.events")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    # Teamly: remove user-created pages (seed pages have id <= 3); ensure a space.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('PROJECTS', 'Проекты',
                        'Пространство для управления портфелем проектов отделов.')
                ON CONFLICT (key) DO NOTHING
            """)
    except Exception as e:
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    # Noise calendar events (RU) that must remain untouched by the agent.
    cur.execute("""INSERT INTO gcal.events (summary, start_datetime, end_datetime, status)
        VALUES ('Ежедневная планёрка', %s, %s, 'confirmed')""",
        (launch_dt.replace(hour=9), launch_dt.replace(hour=9, minute=30)))
    cur.execute("""INSERT INTO gcal.events (summary, start_datetime, end_datetime, status)
        VALUES ('Обеденный перерыв', %s, %s, 'confirmed')""",
        (launch_dt.replace(hour=12), launch_dt.replace(hour=13)))
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    inbox_id = row[0] if row else 1
    cur.execute("""INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
        VALUES (%s, '<noise-001@co.com>', 'Еженедельная рассылка', 'newsletter@company.com', %s, %s, 'Новости этой недели...', true)""",
        (inbox_id, json.dumps(['all@company.com']), launch_dt - timedelta(hours=5)))
    cur.execute("""INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
        VALUES (%s, '<noise-002@co.com>', 'Обслуживание серверов', 'ops@company.com', %s, %s, 'Плановые работы в субботу', false)""",
        (inbox_id, json.dumps(['team@company.com']), launch_dt - timedelta(hours=3)))
    # Noise teamly page (RU) — leftover the agent must ignore; must NOT satisfy
    # the dashboard check.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("SELECT id FROM teamly.spaces WHERE key = 'PROJECTS'")
            srow = cur.fetchone()
            if srow is None:
                cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
                srow = cur.fetchone()
            space_id = srow[0] if srow else None
            if space_id is not None:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) "
                    "VALUES (%s, %s, %s, %s)",
                    (space_id, NOISE_PAGE_TITLE,
                     "Старые заметки со встреч команды. Не относится к текущей задаче.",
                     "team"),
                )
    except Exception as e:
        print(f"[preprocess] WARNING: noise teamly page skipped: {e}")
    conn.commit()
    cur.close()
    conn.close()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)

if __name__ == "__main__":
    main()
