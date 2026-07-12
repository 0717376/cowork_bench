"""Preprocess для rzd-kulinar-team-trip-spb-catering-excel-gcal.

Очищает email/gcal и инжектит контекст:
- 2 события в gcal на 2026-03-10 (утреннее общее собрание в Мск, встреча с СПб-партнёром)
- 1 входящее письмо от events@company.ru с просьбой подготовить план поездки
"""
import argparse
import json
import os
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_tables(conn):
    with conn.cursor() as cur:
        try:
            cur.execute("DELETE FROM email.attachments")
        except Exception:
            pass
        try:
            cur.execute("DELETE FROM email.sent_log")
        except Exception:
            pass
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
        cur.execute("DELETE FROM gcal.events")
    conn.commit()
    print("[preprocess] Очищены email, gcal.")


def get_or_create_inbox(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        conn.commit()
        return cur.fetchone()[0]


def inject_gcal_events(conn):
    events = [
        {
            "id": str(uuid.uuid4()),
            "summary": "Ежемесячное общее собрание команды",
            "start_datetime": "2026-03-10T09:00:00+03:00",
            "end_datetime": "2026-03-10T10:00:00+03:00",
            "start_timezone": "Europe/Moscow",
            "end_timezone": "Europe/Moscow",
            "creator": json.dumps({"email": "admin@company.ru"}),
            "organizer": json.dumps({"email": "admin@company.ru"}),
            "attendees": json.dumps([{"email": "all@company.ru"}]),
            "description": "Ежемесячное общее онлайн-собрание команды. Участники тимбилдинга будут в дороге и пропустят его.",
            "location": "Офис Москва, переговорная А",
        },
        {
            "id": str(uuid.uuid4()),
            "summary": "Встреча с партнёром в Санкт-Петербурге",
            "start_datetime": "2026-03-10T15:00:00+03:00",
            "end_datetime": "2026-03-10T16:00:00+03:00",
            "start_timezone": "Europe/Moscow",
            "end_timezone": "Europe/Moscow",
            "creator": json.dumps({"email": "bd@company.ru"}),
            "organizer": json.dumps({"email": "bd@company.ru"}),
            "attendees": json.dumps([{"email": "events@company.ru"}, {"email": "partner@spb.ru"}]),
            "description": "Встреча по стратегии с партнёром в Санкт-Петербурге. Основной повод тимбилдинг-выезда.",
            "location": "Офис партнёра, Санкт-Петербург",
        },
    ]
    with conn.cursor() as cur:
        for ev in events:
            cur.execute("""
                INSERT INTO gcal.events
                    (id, summary, start_datetime, end_datetime, start_timezone, end_timezone,
                     creator, organizer, attendees, description, location)
                VALUES (%s, %s, %s::timestamptz, %s::timestamptz, %s, %s,
                        %s::jsonb, %s::jsonb, %s::jsonb, %s, %s)
            """, (
                ev["id"], ev["summary"],
                ev["start_datetime"], ev["end_datetime"],
                ev["start_timezone"], ev["end_timezone"],
                ev["creator"], ev["organizer"], ev["attendees"],
                ev["description"], ev["location"],
            ))
    conn.commit()
    print(f"[preprocess] Создано {len(events)} событий в GCal.")


def inject_email(conn, folder_id):
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
            VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s)
        """, (
            folder_id,
            str(uuid.uuid4()),
            "Тимбилдинг-выезд в Санкт-Петербург — нужна координация",
            "events@company.ru",
            json.dumps(["coordinator@company.ru"]),
            (
                "Здравствуйте,\n\nМы планируем тимбилдинг-выезд из Москвы в Санкт-Петербург "
                "10 марта 2026 года для 15 сотрудников. Просим Вас скоординировать "
                "полный план: билеты на «Сапсан», меню ужина в русском стиле "
                "и расписание дня.\n\n"
                "Пожалуйста, подготовьте:\n"
                "1. План переезда (предпочтительно ранний «Сапсан» 06:50)\n"
                "2. Меню кейтеринг-ужина на вечер (около 18:00)\n"
                "3. Полное расписание дня в Excel\n"
                "4. Страницу плана в базе знаний команды\n"
                "5. Все события в общем календаре\n\n"
                "Ориентир бюджета: 7000 ₽ на человека на переезд.\n\n"
                "Пришлите готовый план ответом на это письмо.\n\nСпасибо,\nevents@company.ru"
            ),
        ))
    conn.commit()
    print("[preprocess] Создано входящее письмо от events@company.ru.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_gcal_events(conn)
        folder_id = get_or_create_inbox(conn)
        inject_email(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Подготовка rzd-kulinar-team-trip-spb-catering-excel-gcal завершена.")


if __name__ == "__main__":
    main()
