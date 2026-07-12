"""
Preprocess for terminal-moex-kulinar-excel-word-gcal task.

Очищает Google Calendar. Внедряет существующие события календаря (часть из них
конфликтует с целевыми средами) плюс шумовые события. Данные MOEX и Kulinar
доступны только для чтения (засеяны глобально, инъекция не требуется).
"""
import argparse
import os

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
    conn.commit()
    print("[preprocess] Cleared gcal events.")


def inject_calendar_events(conn):
    """Inject existing calendar events. Two conflict with target Wednesdays,
    plus noise events on other days. All dates are ABSOLUTE (eval checks the
    fixed window 2026-03-16..2026-04-10, so launch_time is ignored)."""

    events = [
        # Conflict: Wed March 18 11:00 (forces session 1 to Thu March 19)
        ('Общее собрание компании', '2026-03-18 11:00:00', '2026-03-18 13:00:00',
         'Ежеквартальное общее собрание для всех сотрудников', 'Главный конференц-зал'),
        # Conflict: Wed April 1 10:00 (forces session 3 to Thu April 2)
        ('Бюджетное ревью за I квартал', '2026-04-01 10:00:00', '2026-04-01 14:00:00',
         'Бюджетное ревью по итогам квартала с финансовым отделом', 'Переговорная совета'),
        # Noise: Tue March 17 09:00
        ('Ежедневный стендап', '2026-03-17 09:00:00', '2026-03-17 09:30:00',
         'Ежедневная планёрка команды', 'Кабинет 201'),
        # Noise: Fri March 20 10:00
        ('Адаптация новых сотрудников (HR)', '2026-03-20 10:00:00', '2026-03-20 12:00:00',
         'Ознакомительная сессия для новых сотрудников', 'Кабинет 105'),
        # Noise: Tue March 24 14:00
        ('Тренинг по продажам', '2026-03-24 14:00:00', '2026-03-24 16:00:00',
         'Ежемесячный тренинг отдела продаж', 'Кабинет 302'),
        # Noise: Sun April 5 22:00
        ('Окно ИТ-обслуживания', '2026-04-05 22:00:00', '2026-04-06 02:00:00',
         'Плановое обслуживание серверов', 'ЦОД'),
    ]

    with conn.cursor() as cur:
        for summary, start, end, desc, loc in events:
            cur.execute(
                "INSERT INTO gcal.events (summary, start_datetime, end_datetime, description, location) "
                "VALUES (%s, %s, %s, %s, %s)",
                (summary, start, end, desc, loc),
            )
    conn.commit()
    print("[preprocess] Injected 6 calendar events (2 conflicts + 4 noise).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_calendar_events(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
