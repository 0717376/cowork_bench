"""
Preprocess for insales-refund-root-cause-excel-word-email task (InSales store).

InSales (wc.*) data is read-only and russified centrally by db/zzz_wc_after_init.sql.
This script:
1. Clears email data
2. Injects 2-3 noise emails (RU) in inbox
"""
import os
import argparse
import json
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def fix_full_refund_amounts(cur):
    """Audit v0.2.3: fully refunded orders (39/49/110 in the shared wc seed)
    must have refund amount equal to the order total. The shared seed is
    read-only, so align it per-task here (fresh PG per task)."""
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
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] Email data cleared.")


def ensure_folders(cur):
    """Ensure INBOX and Sent folders exist."""
    for folder_name in ["INBOX", "Sent"]:
        cur.execute("SELECT id FROM email.folders WHERE name = %s LIMIT 1", (folder_name,))
        if not cur.fetchone():
            cur.execute("INSERT INTO email.folders (name) VALUES (%s)", (folder_name,))
    print("[preprocess] Email folders ensured.")


def inject_noise_emails(cur):
    """Inject a few noise emails so the inbox isn't empty."""
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    inbox_row = cur.fetchone()
    inbox_id = inbox_row[0] if inbox_row else 1

    noise_emails = [
        {
            "subject": "Результаты маркетинговой кампании за I квартал",
            "from_addr": "marketing@company.com",
            "to_addr": json.dumps(["team@company.com"]),
            "body_text": "Привет, команда! Во вложении результаты маркетинговой кампании за I квартал. "
                         "В целом конверсия выросла на 12% по сравнению с прошлым кварталом.",
        },
        {
            "subject": "График обслуживания офиса — март",
            "from_addr": "facilities@company.com",
            "to_addr": json.dumps(["all@company.com"]),
            "body_text": "Обращаем внимание, что обслуживание системы вентиляции запланировано на "
                         "15-16 марта. В этот период температура в здании может колебаться.",
        },
        {
            "subject": "Re: Обед командой в пятницу",
            "from_addr": "sarah.jones@company.com",
            "to_addr": json.dumps(["qa_team@company.com"]),
            "body_text": "Отлично! Забронирую итальянский ресторан. Пожалуйста, подтвердите участие "
                         "до конца четверга.",
        },
    ]

    for email in noise_emails:
        cur.execute("""
            INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr,
                                        date, body_text, is_read, is_important, is_flagged)
            VALUES (%s, %s, %s, %s, %s, NOW(), %s, true, false, false)
        """, (
            inbox_id,
            f"<noise-{uuid.uuid4()}@company.com>",
            email["subject"],
            email["from_addr"],
            email["to_addr"],
            email["body_text"],
        ))

    print(f"[preprocess] Injected {len(noise_emails)} noise emails.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        fix_full_refund_amounts(cur)
        clear_emails(cur)
        ensure_folders(cur)
        inject_noise_emails(cur)
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
