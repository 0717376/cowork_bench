"""
Preprocess for yt-transcript-teamly-song-report-gcal-email task.

Injects:
  - 1 email from editorial@musicblog.com asking for the analysis report
  - 2 gcal events: editorial meeting and content planning session
  - Clears email and gcal tables before injecting

Teamly (RU corporate knowledge base, replaces Notion): we ensure a space exists
for the agent to create the "Afrobeat Mix 2024 - Blog Analysis" page in, and
clear any such analysis page left over from previous runs (idempotency). We do
NOT pre-create the page nor the xlsx — the agent must produce them itself.

Prerequisites:
  - PostgreSQL cowork_gym database running on localhost:5432
  - Teamly schema seeded (db/zzz_teamly_after_init.sql)
"""
import argparse
import json
import os
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

GCAL_EVENTS = [
    {
        "summary": "Ежемесячная редакционная встреча",
        "description": "Ежемесячная общая встреча редакции. Повестка: обзор контент-пайплайна, планирование на II квартал.",
        "start": "2026-03-10 14:00:00",
        "end": "2026-03-10 15:00:00",
    },
    {
        "summary": "Сессия контент-планирования",
        "description": "Квартальная сессия планирования контента. Темы: предстоящие материалы, коллаборации с исполнителями, стратегия в соцсетях.",
        "start": "2026-03-16 10:00:00",
        "end": "2026-03-16 11:00:00",
    },
]


def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Cleared gcal and email tables.")


def inject_gcal_events(conn):
    with conn.cursor() as cur:
        for ev in GCAL_EVENTS:
            cur.execute("""
                INSERT INTO gcal.events (summary, description, start_datetime, end_datetime,
                    start_timezone, end_timezone, creator, organizer, attendees)
                VALUES (%s, %s, %s, %s, %s, %s, %s::jsonb, %s::jsonb, %s::jsonb)
            """, (
                ev["summary"], ev["description"], ev["start"], ev["end"],
                "Europe/Moscow", "Europe/Moscow",
                json.dumps({}), json.dumps({}), json.dumps([]),
            ))
    conn.commit()
    print(f"[preprocess] Injected {len(GCAL_EVENTS)} GCal events.")


def inject_emails(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            conn.commit()
            row = cur.fetchone()
        folder_id = row[0]

        email_data = {
            "message_id": "msg-editorial-001",
            "subject": "Afrobeat Mix Analysis - Status Update Needed",
            "from_addr": "editorial@musicblog.com",
            "to_addr": json.dumps(["editor@afrobeatstoday.com"]),
            "date": "2026-03-07 09:00:00+00",
            "body_text": (
                "Здравствуйте!\n\n"
                "Нам нужен полный аналитический отчёт по видео Afrobeat-микса (7ZQzGq32kAY) к концу недели. "
                "Пожалуйста, включите трек-лист, разбор по исполнителям и план публикаций для трёх запланированных статей. "
                "Убедитесь, что отчёт будет готов к редакционной встрече 10 марта.\n\n"
                "С уважением,\nРедакция\neditorial@musicblog.com"
            ),
            "folder_id": folder_id,
        }
        cur.execute("""
            INSERT INTO email.messages (message_id, subject, from_addr, to_addr, date, body_text, folder_id)
            VALUES (%(message_id)s, %(subject)s, %(from_addr)s, %(to_addr)s::jsonb,
                    %(date)s, %(body_text)s, %(folder_id)s)
            ON CONFLICT (message_id) DO NOTHING
        """, email_data)
    conn.commit()
    print("[preprocess] Injected editorial request email.")


def setup_teamly(conn):
    """Ensure a Teamly space exists for the blog analysis page and clear any
    prior 'Afrobeat Mix 2024 - Blog Analysis' page (idempotency).

    We intentionally do NOT pre-create the page — the agent must create it so
    the evaluation actually exercises the agent's work.
    """
    with conn.cursor() as cur:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is None:
            print("[preprocess] WARNING: teamly schema not found; "
                  "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
            return
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('BLOG', 'Музыкальный блог',
                    'Аналитика музыкальных миксов и материалов блога AfroBeats Today.')
            ON CONFLICT (key) DO NOTHING
        """)
        cur.execute("""
            DELETE FROM teamly.pages
             WHERE title ILIKE '%afrobeat mix%'
                OR title ILIKE '%blog analysis%'
        """)
    conn.commit()
    print("[preprocess] Teamly ready: 'BLOG' space ensured, prior analysis pages cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_gcal_events(conn)
        inject_emails(conn)
        setup_teamly(conn)
    finally:
        conn.close()

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
