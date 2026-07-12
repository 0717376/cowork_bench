"""Preprocess for terminal-insales-pdf-excel-ppt-email.
Clears email. InSales store data is read-only. Injects RU noise emails."""
import argparse
import json
import os
import uuid
import glob as globmod

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Clear email
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

        # Inject noise emails
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            folder_id = row[0]
            cur.execute("SELECT COALESCE(MAX(id), 0) FROM email.messages")
            max_id = cur.fetchone()[0]

            noise_emails = [
                (max_id + 1, "Еженедельная сводка продаж", "sales@company.com",
                 json.dumps(["team@company.com"]),
                 "Прилагаем еженедельную сводку продаж для ознакомления."),
                (max_id + 2, "Уведомление об обслуживании офиса", "facilities@company.com",
                 json.dumps(["all@company.com"]),
                 "В эти выходные будет проводиться обслуживание системы вентиляции офиса."),
                (max_id + 3, "Счёт от поставщика №INV-2026-0315", "accounts@supplier.com",
                 json.dumps(["finance@company.com"]),
                 "Во вложении счёт за недавние поставки."),
                (max_id + 4, "Корпоративный тимбилдинг", "hr@company.com",
                 json.dumps(["team@company.com"]),
                 "Приглашаем вас на ежеквартальное корпоративное мероприятие в следующую пятницу."),
            ]
            for eid, subj, from_addr, to_addr, body in noise_emails:
                cur.execute("""
                    INSERT INTO email.messages (id, folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                    VALUES (%s, %s, %s, %s, %s, %s::jsonb, NOW(), %s, false)
                """, (eid, folder_id, f"noise-{uuid.uuid4()}@company.com",
                      subj, from_addr, to_addr, body))
            print("[preprocess] Injected 4 noise emails.")

        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Clean up agent workspace if provided
    if args.agent_workspace:
        for pattern in ["Recall_Impact_Assessment.xlsx", "Recall_Briefing.pptx",
                        "recall_analysis.py", "customer_impact.py",
                        "recall_impact.json", "customer_impact.json"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
