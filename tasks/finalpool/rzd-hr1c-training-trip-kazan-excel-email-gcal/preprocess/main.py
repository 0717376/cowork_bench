"""
Preprocess for rzd-hr1c-training-trip-kazan-excel-email-gcal.

- Чистит gcal.events и email.* (HR1c — read-only, его не трогаем; rzd — фикс. данные).
- Инжектит:
    1) GCal-событие "Корпоративный тренинг — кикофф" 17.03.2026 13:00-17:00 в Казани.
    2) Входящее письмо от HR-департамента (training@hr.company.ru → hrmanager@company.ru)
       с просьбой организовать командировку.
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
        for table in ("email.attachments", "email.sent_log",
                      "email.messages", "email.drafts"):
            try:
                cur.execute(f"DELETE FROM {table}")
            except Exception:
                pass
    conn.commit()
    print("[preprocess] Cleared email and gcal tables.")


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
                start_timezone, end_timezone, location)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (
            "Корпоративный тренинг — кикофф",
            "Стартовая сессия корпоративного тренинга для отделов Продажи и Маркетинг.",
            "2026-03-17 13:00:00",
            "2026-03-17 17:00:00",
            "Europe/Moscow",
            "Europe/Moscow",
            "Казань, учебный центр на Кремлёвской набережной",
        ))
    conn.commit()
    print("[preprocess] Injected GCal event: Корпоративный тренинг — кикофф.")


def inject_email(conn, folder_id):
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
            VALUES (%s, %s, %s, %s, %s::jsonb, %s, %s)
        """, (
            folder_id,
            "msg-training-hr-001",
            "Организация командировки — тренинг в Казани 17 марта",
            "training@hr.company.ru",
            json.dumps(["hrmanager@company.ru"]),
            "2026-03-13 09:00:00+03",
            (
                "Здравствуйте!\n\n"
                "Нужно организовать корпоративную поездку сотрудников на тренинг в Казань "
                "17 марта 2026 года. Кикофф в 13:00 по местному времени. Прошу подобрать "
                "до пяти сотрудников из отделов Продажи и Маркетинг со стажем не менее "
                "3 лет, забронировать билеты «Стриж» Москва-Казанская ↔ Казань-Пассажирская "
                "(туда и обратно одним днём), подготовить отчёт в Excel с разбивкой по "
                "сотрудникам и бюджету, завести события в общем календаре и отправить нам "
                "сводку, а также подтверждения участникам.\n\n"
                "Спасибо,\nHR-департамент"
            ),
        ))
    conn.commit()
    print("[preprocess] Injected email from training@hr.company.ru.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_gcal(conn)
        folder_id = ensure_email_folder(conn)
        inject_email(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Preprocessing complete for rzd-hr1c-training-trip-kazan-excel-email-gcal.")


if __name__ == "__main__":
    main()
