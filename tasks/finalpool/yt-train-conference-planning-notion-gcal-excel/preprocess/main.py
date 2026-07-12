"""
Preprocess for yt-train-conference-planning-notion-gcal-excel task.

Clears writable tables (email, gcal, gform, teamly) to provide a clean environment.
YouTube and train data are read-only.

Prerequisites:
  - PostgreSQL cowork_gym database running on localhost:5432
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


def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
        # Clear teamly pages created by the agent in the EVENTS space.
        # Seeded teamly pages (ids <= 63 across other spaces) are preserved as
        # format examples; the EVENTS space has no seed pages.
        try:
            cur.execute("""
                DELETE FROM teamly.pages
                WHERE space_id = (SELECT id FROM teamly.spaces WHERE key = 'EVENTS')
            """)
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Cleared writable tables.")


def ensure_email_folder(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            conn.commit()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        ensure_email_folder(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
