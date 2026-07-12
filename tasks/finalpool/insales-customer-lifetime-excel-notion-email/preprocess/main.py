"""
Preprocess script for insales-customer-lifetime-excel-notion-email task.

InSales (InSales) data is read-only. This script:
1. Ensures an EMPTY Teamly space 'Customer CRM' exists (clears prior CRM pages
   idempotently; does NOT pre-seed the 50 customer pages the agent must create).
2. Clears email data and injects noise emails.
"""

import os
import argparse
import json
import uuid
from datetime import datetime, timezone

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# Teamly space the agent must populate with one page per customer.
SPACE_KEY = "CUSTCRM"
SPACE_NAME = "Customer CRM"
SPACE_DESC = "CRM knowledge base: one page per customer with CLV tier and at-risk status."


def setup_teamly(cur):
    """Ensure an EMPTY 'Customer CRM' Teamly space exists.

    Global seed pages have id <= 3 (zzz_teamly_after_init.sql). We create the
    space idempotently and remove any prior CRM pages (user-created, id > 3)
    so reruns start clean. The 50 customer answer-pages are NOT pre-seeded.
    """
    print("[preprocess] Setting up Teamly 'Customer CRM' space...")

    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly.spaces not found; skipping Teamly setup.")
        return

    cur.execute(
        """
        INSERT INTO teamly.spaces (key, name, description)
        VALUES (%s, %s, %s)
        ON CONFLICT (key) DO UPDATE SET name = EXCLUDED.name,
                                        description = EXCLUDED.description
        """,
        (SPACE_KEY, SPACE_NAME, SPACE_DESC),
    )

    # Idempotently clear prior user-created pages in the CRM space (id > 3 keeps
    # the global format-example seeds intact).
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute(
            """
            DELETE FROM teamly.pages
            WHERE id > 3
              AND space_id IN (SELECT id FROM teamly.spaces WHERE key = %s)
            """,
            (SPACE_KEY,),
        )
    print("[preprocess] Teamly 'Customer CRM' space ready (empty).")


def clear_emails(cur):
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def inject_email_noise(cur):
    print("[preprocess] Injecting noise emails...")
    now = datetime.now(timezone.utc)

    # Get INBOX folder id
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if not row:
        print("[preprocess] WARNING: No INBOX folder found, skipping email noise.")
        return
    inbox_id = row[0]

    noise_emails = [
        {
            "subject": "Monthly Newsletter - March 2026",
            "from": "newsletter@marketing.com",
            "to": ["admin@company.com"],
            "body": "Here is your monthly marketing newsletter with the latest updates and promotions.",
        },
        {
            "subject": "Server Maintenance Scheduled",
            "from": "devops@company.com",
            "to": ["admin@company.com"],
            "body": "Planned maintenance window this weekend. Expect 2 hours of downtime Saturday night.",
        },
    ]

    for em in noise_emails:
        cur.execute(
            """INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date,
               body_text, is_read, is_important, is_flagged)
               VALUES (%s, %s, %s, %s, %s, %s, %s, false, false, false)""",
            (
                inbox_id,
                f"<{uuid.uuid4()}@noise.com>",
                em["subject"],
                em["from"],
                json.dumps(em["to"]),
                now,
                em["body"],
            ),
        )
    print(f"[preprocess] Injected {len(noise_emails)} noise emails.")


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
        inject_email_noise(cur)
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
