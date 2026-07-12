"""Preprocess: clear user-created Teamly pages for a clean state.

Seed pages (id <= 3) are kept as a format example. No answers are pre-seeded.
"""
import os
import argparse
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            # Drop user-created pages; keep seed pages (id <= 3) as format example.
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        conn.commit()
        print("Cleared user-created Teamly pages (id > 3)")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
