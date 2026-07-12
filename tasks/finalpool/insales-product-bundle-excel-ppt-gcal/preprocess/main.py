"""
Preprocess script for insales-product-bundle-excel-ppt-gcal task.
Clears Google Calendar and email data, injects noise calendar events.
InSales (wc.*) data is read-only.
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
        # Clear Google Calendar events
        cur.execute("DELETE FROM gcal.events")
        print("[preprocess] Cleared Google Calendar events.")

        # Clear email data
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

        # Inject noise calendar events
        noise_events = [
            ("noise-event-001", "Ежедневный стендап", "Ежедневная планёрка команды",
             "2026-03-16 09:00:00+00", "2026-03-16 09:15:00+00"),
            ("noise-event-002", "Обед с поставщиком", "Обсуждение цепочки поставок Q2",
             "2026-03-17 12:00:00+00", "2026-03-17 13:00:00+00"),
            ("noise-event-003", "Синхронизация маркетинга", "Обзор метрик кампаний",
             "2026-03-19 14:00:00+00", "2026-03-19 15:00:00+00"),
        ]
        for eid, summary, desc, start, end in noise_events:
            cur.execute(
                """INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime)
                   VALUES (%s, %s, %s, %s, %s) ON CONFLICT (id) DO NOTHING""",
                (eid, summary, desc, start, end)
            )
        print("[preprocess] Injected 3 noise calendar events.")

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
