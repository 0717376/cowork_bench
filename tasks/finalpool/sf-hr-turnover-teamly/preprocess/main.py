"""Preprocess: clear email and Teamly data, ensure an empty HR space exists.

Idempotent. Does NOT pre-seed the deliverable page/table — only clears leftovers
and makes sure a Teamly space is available for the agent to write into.
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

SPACE_KEY = "HR"
SPACE_NAME = "Управление персоналом"


def clear_emails(cur):
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.drafts")
    cur.execute("DELETE FROM email.messages")
    print("[preprocess] Email data cleared.")


def clear_teamly(cur):
    """Clear Teamly pages/labels idempotently (leave nothing for the agent to copy)."""
    print("[preprocess] Clearing Teamly data...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; skipping.")
        return
    cur.execute("DELETE FROM teamly.page_labels")
    cur.execute("DELETE FROM teamly.labels")
    cur.execute("DELETE FROM teamly.pages")
    print("[preprocess] Teamly pages cleared.")


def ensure_space(cur):
    """Ensure an empty HR space exists for the agent to write the page into."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        return
    cur.execute(
        """INSERT INTO teamly.spaces (key, name, description)
           VALUES (%s, %s, %s) ON CONFLICT (key) DO NOTHING""",
        (SPACE_KEY, SPACE_NAME, "Аналитика по персоналу, отчёты по отделам и оплате труда."),
    )
    print(f"[preprocess] Ensured Teamly space '{SPACE_KEY}'.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = True
    cur = conn.cursor()

    clear_emails(cur)
    clear_teamly(cur)
    ensure_space(cur)

    cur.close()
    conn.close()
    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
