"""Preprocess for rzd-msk-spb-team-trip-gcal-word (hard version).

Очищает gcal/notion/email и заполняет контекст:
- 1 событие в календаре: "Встреча с клиентом в Санкт-Петербурге" 2026-03-10 14:00-16:00
- 1 входящее письмо от travel@consulting.ru — поездка для команды из 3 человек,
  бюджет общий, без указания нужного класса
"""
import argparse
import json
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
        cur.execute("DELETE FROM notion.pages")
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
    conn.commit()
    print("[preprocess] Очищены таблицы gcal, notion, email.")


def ensure_email_folder(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        fid = cur.fetchone()[0]
    conn.commit()
    return fid


def inject_gcal(conn):
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO gcal.events (summary, description, start_datetime, end_datetime,
                start_timezone, end_timezone)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            "Встреча с клиентом в Санкт-Петербурге",
            "Стратегическая встреча с ключевым клиентом в офисе СПб. Просьба организовать дорогу.",
            "2026-03-10 14:00:00",
            "2026-03-10 16:00:00",
            "Europe/Moscow",
            "Europe/Moscow",
        ))
    conn.commit()
    print("[preprocess] Создано событие в GCal: Встреча с клиентом.")


def inject_email(conn, folder_id):
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
            VALUES (%s, %s, %s, %s, %s::jsonb, %s, %s)
        """, (
            folder_id,
            "msg-travel-rzd-team-001",
            "Помощь с организацией командной поездки — Москва — СПб 10 марта",
            "travel@consulting.ru",
            json.dumps(["consultant@company.ru"]),
            "2026-03-07 09:00:00+03",
            (
                "Здравствуйте,\n\n10 марта 2026 года у нас запланирована важная встреча с клиентом "
                "в Санкт-Петербурге. Едем командой из трёх человек. Параметры (размер команды и "
                "общий бюджет) указаны в файле trip_config.json в рабочей директории. "
                "Просьба организовать поездку «Сапсаном» туда и обратно в один день — нужно "
                "успеть к встрече с разумным запасом и вернуться в Москву в тот же день. "
                "Класс выбери сам исходя из бюджета — нам важно уложиться в указанную сумму "
                "на всю команду; класс особо не важен, главное чтобы все ехали в одном "
                "и не было превышения. Если бюджета не хватит даже на самый дешёвый класс — "
                "сообщи об этом отдельно, без публикации плана.\n\nПожалуйста, оформи "
                "официальный план поездки в виде документа.\n\nС уважением,\n"
                "Отдел организации поездок\ntravel@consulting.ru"
            ),
        ))
    conn.commit()
    print("[preprocess] Создано входящее письмо от travel@consulting.ru.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_gcal(conn)
        folder_id = ensure_email_folder(conn)
        inject_email(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Подготовка для rzd-msk-spb-trip-notion-gcal-word завершена.")


if __name__ == "__main__":
    main()
