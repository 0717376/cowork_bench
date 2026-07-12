"""
Preprocess script for canvas-quiz-item-analysis-word-gcal-email task.
Clears gcal and email data, injects noise events and emails.
Canvas data is read-only.
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        # Clear gcal events
        cur.execute("DELETE FROM gcal.events")
        print("[preprocess] Cleared gcal events.")

        # Clear email data
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

        # Inject 3 noise calendar events
        noise_events = [
            {
                "id": str(uuid.uuid4()),
                "summary": "Совещание кафедры",
                "description": "Ежемесячная синхронизация кафедры",
                "start_datetime": "2026-03-10 09:00:00+00",
                "end_datetime": "2026-03-10 10:00:00+00",
                "status": "confirmed",
            },
            {
                "id": str(uuid.uuid4()),
                "summary": "Обеденный перерыв преподавателей",
                "description": "Неформальная встреча преподавателей",
                "start_datetime": "2026-03-12 12:00:00+00",
                "end_datetime": "2026-03-12 13:00:00+00",
                "status": "confirmed",
            },
            {
                "id": str(uuid.uuid4()),
                "summary": "Пересмотр бюджета",
                "description": "Сессия планирования бюджета на II квартал",
                "start_datetime": "2026-03-14 14:00:00+00",
                "end_datetime": "2026-03-14 15:30:00+00",
                "status": "confirmed",
            },
        ]
        for evt in noise_events:
            cur.execute(
                """INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status)
                   VALUES (%s, %s, %s, %s, %s, %s)""",
                (evt["id"], evt["summary"], evt["description"],
                 evt["start_datetime"], evt["end_datetime"], evt["status"]),
            )
        print(f"[preprocess] Injected {len(noise_events)} noise calendar events.")

        # Inject 2 noise emails
        cur.execute("SELECT id FROM email.folders WHERE name = 'Inbox' LIMIT 1")
        row = cur.fetchone()
        if row:
            inbox_id = row[0]
        else:
            cur.execute("INSERT INTO email.folders (name) VALUES ('Inbox') RETURNING id")
            inbox_id = cur.fetchone()[0]

        noise_emails = [
            {
                "subject": "Напоминание: срок сдачи итоговых оценок",
                "from_addr": "registrar@university.edu",
                "to_addr": json.dumps(["faculty@university.edu"]),
                "body_text": "Пожалуйста, выставите все итоговые оценки до 20 марта 2026 года.",
            },
            {
                "subject": "Обновление по парковке кампуса",
                "from_addr": "facilities@university.edu",
                "to_addr": json.dumps(["all-staff@university.edu"]),
                "body_text": "Парковка B будет закрыта на обслуживание на следующей неделе.",
            },
        ]
        for em in noise_emails:
            cur.execute(
                """INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, body_text)
                   VALUES (%s, %s, %s, %s, %s, %s)""",
                (inbox_id, str(uuid.uuid4()), em["subject"],
                 em["from_addr"], em["to_addr"], em["body_text"]),
            )
        print(f"[preprocess] Injected {len(noise_emails)} noise emails.")

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
