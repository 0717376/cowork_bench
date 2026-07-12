"""Preprocess for canvas-module-completion-teamly-gcal.

- Ensures a Teamly space exists for the agent to create the module-tracker page
  in, and clears any prior tracker pages (idempotency).
- Clears Google Calendar events and email data idempotently.

We intentionally do NOT pre-create the tracker page, the calendar events, nor
the email — the agent must produce them itself so the evaluation actually tests
the agent. Canvas course 7 (the seeded source of truth) is left untouched.
"""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def setup_teamly(cur):
    """Ensure the course Teamly space exists and clear prior tracker pages."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return
    # Dedicated space for the agent to drop the module-tracker page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('CCC', 'Creative Computing and Culture',
                'Рабочее пространство преподавательской команды курса '
                'Creative Computing and Culture (Fall 2014): обзоры модулей и учёт прогресса.')
        ON CONFLICT (key) DO NOTHING
    """)
    # Idempotency: drop any tracker pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%module tracker%'
            OR title ILIKE '%ccc fall 2014%'
            OR title ILIKE '%трекер%модул%'
    """)
    print("[preprocess] Teamly ready: 'CCC' space ensured, prior tracker pages cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        setup_teamly(cur)
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared GCal and email data. Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
