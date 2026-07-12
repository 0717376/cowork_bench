"""
Preprocess for kulinar-nutrition-excel-email task (russified -> kulinar).
- Clear email data so agent starts fresh.
- kulinar is a standalone MCP server (recipes seeded in JSON), no DB cleanup needed.
- No answer pre-seeding: only clears email tables for idempotency.
"""
import os
import argparse
import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_email(conn):
    """Clear all email data."""
    with conn.cursor() as cur:
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        cur.execute("DELETE FROM email.drafts")
    conn.commit()
    print("[preprocess] Cleared email data")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        clear_email(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
