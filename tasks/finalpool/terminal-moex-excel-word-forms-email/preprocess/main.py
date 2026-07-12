"""
Preprocess for terminal-moex-excel-word-forms-email task.

Clears gform and email tables. Injects noise data (Russian).
MOEX Finance (moex.* schema) is read-only and globally seeded.
"""
import argparse
import json
import os
from datetime import datetime, timedelta

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            conn.rollback()
    conn.commit()
    print("[preprocess] Cleared gform and email tables.")


def inject_noise_gform(conn):
    """Inject noise form data the agent should ignore."""
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO gform.forms (id, title, document_title, description)
            VALUES ('noise_form_001', 'Заявка на офисные принадлежности',
                    'Заявка на офисные принадлежности',
                    'Используйте эту форму для заказа офисных принадлежностей.')
        """)
        cur.execute("""
            INSERT INTO gform.questions (id, form_id, item_id, title, question_type, required, config, position)
            VALUES ('noise_q_001', 'noise_form_001', 'noise_item_001',
                    'Какие принадлежности вам нужны?', 'PARAGRAPH', true, '{}'::jsonb, 0)
        """)
    conn.commit()
    print("[preprocess] Injected noise gform data.")


def inject_noise_emails(conn, launch_dt):
    """Inject noise emails."""
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            row = cur.fetchone()
            conn.commit()
        folder_id = row[0]
        d1 = (launch_dt - timedelta(days=56)).strftime("%Y-%m-%d %H:%M:%S")
        d2 = (launch_dt - timedelta(days=54)).strftime("%Y-%m-%d %H:%M:%S")
        cur.execute(f"""
            INSERT INTO email.messages (folder_id, subject, from_addr, to_addr, body_text, date)
            VALUES
            (%s, 'Обзор рынка: еженедельная сводка', 'market_updates@broker.com',
             '["portfolio_team@company.com"]'::jsonb,
             'На этой неделе рынок был разнонаправленным, рост возглавил технологический сектор.', '{d1}'),
            (%s, 'Напоминание о заседании совета директоров', 'admin@company.com',
             '["board@company.com"]'::jsonb,
             'Напоминание: ежеквартальное заседание совета директоров в следующий вторник в 14:00.', '{d2}')
        """, (folder_id, folder_id))
    conn.commit()
    print("[preprocess] Injected noise email data.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    launch_dt = datetime.strptime(args.launch_time, "%Y-%m-%d %H:%M:%S") if args.launch_time else datetime(2026, 3, 7, 10, 0, 0)

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_noise_gform(conn)
        inject_noise_emails(conn, launch_dt)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
