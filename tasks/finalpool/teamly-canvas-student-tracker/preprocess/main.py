"""
Preprocess script for the canvas-student-tracker task (Teamly variant).

Canvas is read-only (source of truth), so no changes there.
This script:
1. Ensures a Teamly space exists for the agent to drop the at-risk page into,
   and clears any prior at-risk pages (idempotency). We do NOT pre-create the
   at-risk page itself — the agent must produce it.
2. Clears email data (messages, attachments, sent_log, drafts)
3. Clears Google Calendar events
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


def setup_teamly(cur):
    """Ensure the academic-affairs Teamly space exists and clear prior at-risk pages."""
    print("[preprocess] Setting up Teamly...")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return
    # Dedicated space for the agent to drop the at-risk students page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('ACADEMIC', 'Учебная часть',
                'Мониторинг успеваемости студентов и сводки по группам риска.')
        ON CONFLICT (key) DO NOTHING
    """)
    # Idempotency: drop any at-risk pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%fff-2014j%'
            OR title ILIKE '%at-risk%'
            OR title ILIKE '%at risk%'
            OR title ILIKE '%группа риска%'
    """)
    print("[preprocess] Teamly ready: 'ACADEMIC' space ensured, prior at-risk pages cleared.")


def clear_emails(cur):
    """Clear all email data except folder structure and account config."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def clear_gcal(cur):
    """Clear all Google Calendar events."""
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] Google Calendar events cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        setup_teamly(cur)
        clear_emails(cur)
        clear_gcal(cur)
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
