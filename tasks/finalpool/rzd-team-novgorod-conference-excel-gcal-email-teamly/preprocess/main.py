"""
Preprocess for rzd-team-novgorod-conference-excel-gcal-email-teamly.

Injects:
- 1 gcal event: "Конференция по древнерусской истории" 2026-03-12 → 2026-03-15
- 2 emails: from moscow_team@uni.ru and spb_team@uni.ru asking about travel
- Clears email, gcal tables. Teamly pages чистим только пользовательские
  (оставляем сид-страницы как пример формата).
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
        # Сидовые страницы teamly оставляем, удаляем только всё после максимального
        # сид-id (3 страницы засеяно в db/zzz_teamly_after_init.sql).
        try:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        except Exception:
            pass
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
    print("[preprocess] Очищены email, gcal, teamly.")


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
            "Конференция по древнерусской истории",
            "Ежегодная научная конференция по истории Древней Руси. "
            "Место: Новгородский государственный университет, конференц-зал.",
            "2026-03-12 09:00:00",
            "2026-03-15 17:00:00",
            "Europe/Moscow",
            "Europe/Moscow",
        ))
    conn.commit()
    print("[preprocess] Создано 1 событие в GCal: Конференция по древнерусской истории.")


def inject_emails(conn, folder_id):
    emails = [
        {
            "message_id": "msg-moscow-team-001",
            "subject": "Поездка на конференцию в Великий Новгород — московская группа",
            "from_addr": "moscow_team@uni.ru",
            "to_addr": ["conference@medieval-rus.ru"],
            "date": "2026-03-05 10:00:00+03",
            "body": (
                "Здравствуйте!\n\nНас трое из московской группы (проф. Иванов, "
                "к.и.н. Соколова, м.н.с. Кузнецова) — едем на конференцию по "
                "древнерусской истории в Великий Новгород 12–15 марта 2026 года. "
                "Просьба организовать ж/д переезд из Москвы 12 марта. Желательно "
                "согласовать прибытие с питерской группой для общего трансфера до "
                "места проведения. Жду расписание и номер поезда.\n\nС уважением,\n"
                "Московская исследовательская группа"
            ),
        },
        {
            "message_id": "msg-spb-team-001",
            "subject": "Поездка на конференцию в Великий Новгород — петербургская группа",
            "from_addr": "spb_team@uni.ru",
            "to_addr": ["conference@medieval-rus.ru"],
            "date": "2026-03-05 11:00:00+03",
            "body": (
                "Здравствуйте!\n\nНас двое из петербургской группы (проф. Петров и "
                "к.и.н. Морозова) — едем на конференцию по древнерусской истории в "
                "Великий Новгород 12–15 марта 2026 года. Просьба согласовать прибытие "
                "с московской группой, чтобы вместе доехать от вокзала до места "
                "проведения. Помогите подобрать поезд из СПб на 12 марта.\n\n"
                "С уважением,\nПетербургская исследовательская группа"
            ),
        },
    ]
    with conn.cursor() as cur:
        for e in emails:
            cur.execute("""
                INSERT INTO email.messages
                    (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
                VALUES (%s, %s, %s, %s, %s::jsonb, %s, %s)
            """, (
                folder_id, e["message_id"], e["subject"], e["from_addr"],
                json.dumps(e["to_addr"]), e["date"], e["body"],
            ))
    conn.commit()
    print("[preprocess] Созданы 2 входящих письма от групп.")


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
        inject_emails(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Подготовка rzd-team-novgorod-conference-excel-gcal-email-teamly завершена.")


if __name__ == "__main__":
    main()
