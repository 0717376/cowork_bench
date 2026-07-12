"""
Preprocess script for quarterly-sales-restock-review task.

Prerequisites:
- PostgreSQL running at localhost:5432 with cowork_gym database
- ClickHouse (sf_data) and InSales (wc) schemas are read-only and pre-populated
- Email and Teamly schemas are writable

This script:
1. Clears email data (messages, attachments, sent_log)
2. Clears Teamly user-created pages (seed pages with id <= 3 are kept)
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


def clear_emails(cur):
    """Clear all email data except folder structure."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def clear_teamly(cur):
    """Clear user-created Teamly pages idempotently (seed pages have id <= 3)."""
    print("[preprocess] Clearing Teamly user pages...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    print("[preprocess] Teamly user pages cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_emails(cur)
        clear_teamly(cur)
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
