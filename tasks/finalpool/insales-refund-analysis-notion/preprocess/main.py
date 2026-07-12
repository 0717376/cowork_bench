"""
Preprocess script for insales-refund-analysis-notion (InSales store + Teamly).

The InSales store data (PG schema wc.*) is read-only and seeded globally.
This script:
1. Clears user-created Teamly pages (seed pages have id <= 3); ensures a space.
2. Clears email data.
It does NOT pre-seed any answer the agent must produce.
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


def clear_teamly(cur):
    print("[preprocess] Clearing Teamly user pages...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute(
            """
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('FINANCE', 'Финансы',
                    'Рабочее пространство финансового отдела: отчётность и аналитика.')
            ON CONFLICT (key) DO NOTHING
            """
        )
    print("[preprocess] Teamly cleared.")


def fix_full_refund_amounts(cur):
    # Audit v0.2.3: fully refunded orders (39/49/110 in the shared wc seed)
    # must have refund amount equal to the order total. The shared seed is
    # read-only, so align it per-task here (fresh PG per task).
    print("[preprocess] Aligning full-refund amounts with order totals...")
    cur.execute(
        """
        UPDATE wc.refunds r
        SET amount = o.total
        FROM wc.orders o
        WHERE r.order_id = o.id
          AND o.status = 'refunded'
          AND r.order_id IN (39, 49, 110)
          AND (SELECT COUNT(*) FROM wc.refunds r2 WHERE r2.order_id = r.order_id) = 1
          AND r.amount IS DISTINCT FROM o.total
        """
    )
    print(f"[preprocess] Full-refund amounts aligned ({cur.rowcount} rows).")


def clear_emails(cur):
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        fix_full_refund_amounts(cur)
        clear_teamly(cur)
        clear_emails(cur)
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
