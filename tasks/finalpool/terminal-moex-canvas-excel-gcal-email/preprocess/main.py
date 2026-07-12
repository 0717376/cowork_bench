"""
Preprocess for terminal-moex-canvas-excel-gcal-email task.

Clears gcal and email tables. Injects conflicting calendar events and noise emails.
Repairs the score scale in canvas.submissions for courses 16/17 (seed stores
percent-scale scores regardless of points_possible). MOEX Finance is read-only.
"""
import argparse
import json
import os
import uuid
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
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            conn.rollback()
    conn.commit()
    print("[preprocess] Cleared gcal and email tables.")


def fix_canvas_score_scale(conn):
    """Rescale percent-scale seed scores to each assignment's points_possible."""
    with conn.cursor() as cur:
        cur.execute("""
            UPDATE canvas.submissions s
            SET score = s.score * a.points_possible / 100
            FROM canvas.assignments a
            WHERE s.assignment_id = a.id
              AND a.course_id IN (16, 17)
              AND a.points_possible NOT IN (0, 100)
              AND s.score IS NOT NULL
        """)
        updated = cur.rowcount
    conn.commit()
    print(f"[preprocess] Rescaled {updated} canvas submission scores for courses 16/17.")


def inject_calendar_conflicts(conn, launch_time):
    """Inject some calendar events that create conflicts on certain days."""
    base = datetime.fromisoformat(launch_time.replace("Z", "+00:00")) if launch_time else datetime.now()
    # Find the next Monday from base
    days_to_monday = (7 - base.weekday()) % 7
    if days_to_monday == 0 and base.hour >= 12:
        days_to_monday = 7
    monday = base + timedelta(days=days_to_monday)
    monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)

    with conn.cursor() as cur:
        # Conflict on first Monday at 2pm (forces workshop 1 to Tuesday)
        conflict_start = monday.replace(hour=14, minute=0)
        conflict_end = monday.replace(hour=16, minute=0)
        cur.execute("""
            INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime,
                                     start_timezone, end_timezone, status)
            VALUES (%s, 'Заседание учёного совета', 'Ежемесячное заседание учёного совета',
                    %s, %s, 'Europe/Moscow', 'Europe/Moscow', 'confirmed')
        """, (str(uuid.uuid4()), conflict_start.isoformat(), conflict_end.isoformat()))

        # Conflict on first Wednesday at 1:30pm (forces workshop 3 to Thursday)
        wed = monday + timedelta(days=2)
        conflict_start2 = wed.replace(hour=13, minute=30)
        conflict_end2 = wed.replace(hour=15, minute=0)
        cur.execute("""
            INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime,
                                     start_timezone, end_timezone, status)
            VALUES (%s, 'Бюджетный комитет кафедры', 'Квартальное заседание бюджетного комитета',
                    %s, %s, 'Europe/Moscow', 'Europe/Moscow', 'confirmed')
        """, (str(uuid.uuid4()), conflict_start2.isoformat(), conflict_end2.isoformat()))

        # Non-conflicting event (morning, should not interfere)
        thu = monday + timedelta(days=3)
        cur.execute("""
            INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime,
                                     start_timezone, end_timezone, status)
            VALUES (%s, 'Приёмные часы', 'Обычные утренние приёмные часы',
                    %s, %s, 'Europe/Moscow', 'Europe/Moscow', 'confirmed')
        """, (str(uuid.uuid4()),
              thu.replace(hour=9, minute=0).isoformat(),
              thu.replace(hour=11, minute=0).isoformat()))

        # Another non-conflicting event on second week
        next_mon = monday + timedelta(days=7)
        cur.execute("""
            INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime,
                                     start_timezone, end_timezone, status)
            VALUES (%s, 'Научный семинар', 'Еженедельный научный семинар',
                    %s, %s, 'Europe/Moscow', 'Europe/Moscow', 'confirmed')
        """, (str(uuid.uuid4()),
              next_mon.replace(hour=10, minute=0).isoformat(),
              next_mon.replace(hour=12, minute=0).isoformat()))

    conn.commit()
    print(f"[preprocess] Injected calendar conflicts (Monday {monday.date()}, Wednesday {wed.date()}).")


def inject_noise_emails(conn):
    """Inject noise emails the agent should ignore."""
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            row = cur.fetchone()
            conn.commit()
        folder_id = row[0]

        # Ensure Sent folder exists
        cur.execute("SELECT id FROM email.folders WHERE name = 'Sent' LIMIT 1")
        sent_row = cur.fetchone()
        if not sent_row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('Sent') RETURNING id")
            sent_row = cur.fetchone()
            conn.commit()

        cur.execute("""
            INSERT INTO email.messages (folder_id, subject, from_addr, to_addr, body_text, date)
            VALUES
            (%s, 'Подтверждение заказа учебников', 'bookstore@university.edu',
             '["prof.chen@university.edu"]'::jsonb,
             'Ваш заказ на 30 экземпляров учебника «Финансовые рынки» Мишкина подтверждён.', '2026-02-28 09:00:00'),
            (%s, 'Расписание весенних каникул', 'registrar@university.edu',
             '["all_faculty@university.edu"]'::jsonb,
             'Весенние каникулы пройдут с 20 по 28 марта. Все занятия приостановлены.', '2026-03-01 10:00:00'),
            (%s, 'Продление парковочного пропуска', 'parking@university.edu',
             '["prof.chen@university.edu"]'::jsonb,
             'Срок действия вашего парковочного пропуска истекает 1 апреля. Пожалуйста, продлите его онлайн.', '2026-03-02 08:30:00'),
            (%s, 'Обновление по научному гранту', 'grants@university.edu',
             '["prof.chen@university.edu"]'::jsonb,
             'Рассмотрение вашей заявки на грант запланировано на следующий месяц.', '2026-03-03 14:00:00')
        """, (folder_id, folder_id, folder_id, folder_id))
    conn.commit()
    print("[preprocess] Injected noise email data.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    launch_time = args.launch_time or "2026-03-07T09:00:00"

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        fix_canvas_score_scale(conn)
        inject_calendar_conflicts(conn, launch_time)
        inject_noise_emails(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
