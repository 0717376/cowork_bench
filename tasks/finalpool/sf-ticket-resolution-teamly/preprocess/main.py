"""
Preprocess for sf-ticket-resolution (ClickHouse + Teamly).
Clears user-created Teamly pages and email tables. ClickHouse sf_data is read-only.
Seed Teamly pages have id <= 3; we only remove agent-created pages (id > 3).
"""
import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        print("[preprocess] Clearing Teamly user data...")
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.page_labels WHERE page_id > 3")
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        # Remove any leftover agent-created space from prior runs (seed keys: TEAM, TRIPS).
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.spaces WHERE key NOT IN ('TEAM', 'TRIPS')")
        print("[preprocess] Teamly user data cleared.")

        print("[preprocess] Clearing email data...")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Email data cleared.")

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
