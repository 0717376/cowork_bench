"""
Preprocess script for moex-market-summary-teamly task.

MOEX Finance is read-only. This script clears Teamly data so the agent
starts from an empty notes space (no answer is pre-seeded).
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        print("[preprocess] Clearing Teamly data...")
        cur.execute("DELETE FROM teamly.page_labels")
        cur.execute("DELETE FROM teamly.labels")
        cur.execute("DELETE FROM teamly.pages")
        cur.execute("DELETE FROM teamly.spaces")
        print("[preprocess] Teamly data cleared.")
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
