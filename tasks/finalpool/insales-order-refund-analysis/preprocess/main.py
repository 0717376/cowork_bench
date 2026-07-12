"""Preprocess: clear email tables for clean state."""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    # Audit v0.2.3: fully refunded orders (39/49/110 in the shared wc seed)
    # must have refund amount equal to the order total. The shared seed is
    # read-only, so align it per-task here (fresh PG per task).
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
    print(f"Full-refund amounts aligned with order totals ({cur.rowcount} rows).")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages WHERE folder_id != 0")
    conn.commit()
    cur.close()
    conn.close()
    print("Email tables cleared for clean state.")


if __name__ == "__main__":
    main()
