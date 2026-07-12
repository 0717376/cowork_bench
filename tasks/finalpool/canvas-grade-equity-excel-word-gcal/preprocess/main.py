"""
Preprocess script for canvas-grade-equity-excel-word-gcal task.
Clears Google Calendar and email data, injects noise calendar events.
Canvas data is read-only.
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

        # Inject 3 noise calendar events (RU titles/locations, not eval-checked)
        noise_events = [
            ("Совещание кафедры", "Еженедельное общее совещание сотрудников",
             "2026-03-16 09:00:00", "2026-03-16 10:00:00", "Аудитория А"),
            ("Обзор бюджета за I квартал", "Квартальный обзор бюджета с финансовым отделом",
             "2026-03-17 14:00:00", "2026-03-17 15:30:00", "Зал заседаний"),
            ("Экскурсия по корпусу", "Осмотр нового учебного корпуса",
             "2026-03-19 11:00:00", "2026-03-19 12:00:00", "Главный холл"),
        ]
        for summary, desc, start, end, location in noise_events:
            cur.execute("""
                INSERT INTO gcal.events (summary, description, start_datetime, end_datetime, location)
                VALUES (%s, %s, %s, %s, %s)
            """, (summary, desc, start, end, location))
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
