"""
Preprocess script for gcal-clickhouse-deadline-tracker task.

ClickHouse (sf_data) data is read-only and globally seeded. This script:
1. Clears Google Calendar events
2. Clears email data
3. Clears Teamly data (user-created pages; ensures a space)
"""

import os
import argparse

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


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
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def clear_teamly(cur):
    """Clear user-created Teamly pages and ensure a knowledge-base space.

    Seed pages have id <= 3 and must be preserved. We only drop user/agent
    pages (id > 3) so the run is idempotent without pre-seeding the answer.
    """
    print("[preprocess] Clearing Teamly data...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('SUPPORT', 'Поддержка',
                    'База знаний команды поддержки: аудиты соблюдения SLA и отчёты.')
            ON CONFLICT (key) DO NOTHING
        """)
    print("[preprocess] Teamly data cleared.")


def clear_gsheet(cur):
    """Clear Google Sheets data."""
    print("[preprocess] Clearing Google Sheets data...")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.sheets")
    try:
        cur.execute("DELETE FROM gsheet.permissions")
    except Exception:
        pass
    cur.execute("DELETE FROM gsheet.spreadsheets")
    print("[preprocess] Google Sheets data cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gcal(cur)
        clear_emails(cur)
        clear_teamly(cur)
        clear_gsheet(cur)
        conn.commit()
        print("[preprocess] Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
