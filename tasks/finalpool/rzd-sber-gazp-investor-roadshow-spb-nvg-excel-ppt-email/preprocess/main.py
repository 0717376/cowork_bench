"""
Preprocess for rzd-sber-gazp-investor-roadshow-spb-nvg-excel-ppt-email.

Clears email tables and injects 2 RU inbox emails:
- from investors@fundmanager.ru (запрос расписания роадшоу)
- from spb_partners@finance.ru (подтверждение доступности в СПб 10.03)
"""
import argparse
import json
import os
import uuid
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
        cur.execute("DELETE FROM email.attachments")
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
    print("[preprocess] Cleared email tables.")


def get_or_create_inbox(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        conn.commit()
        return cur.fetchone()[0]


def inject_emails(conn, folder_id):
    emails = [
        {
            "subject": "Q1 2026 Roadshow — запрос расписания",
            "from_addr": "investors@fundmanager.ru",
            "to_addr": json.dumps(["analyst@firm.ru"]),
            "body": (
                "Здравствуйте,\n\n"
                "Мы заинтересованы в участии в investor roadshow Q1 2026. Пришлите, пожалуйста, "
                "подтверждённое расписание: даты поездок, время прибытия в каждый город, повестку встреч. "
                "Также будем благодарны, если вы заранее предоставите финансовую презентацию и "
                "аналитические материалы — нашей команде нужно подготовить вопросы.\n\n"
                "Особенно интересуют тренды по выручке и EPS за последние два года, "
                "а также прогноз по руководству эмитента.\n\n"
                "С уважением,\ninvestors@fundmanager.ru"
            ),
        },
        {
            "subject": "Подтверждаем доступность — встреча 10 марта",
            "from_addr": "spb_partners@finance.ru",
            "to_addr": json.dumps(["analyst@firm.ru"]),
            "body": (
                "Здравствуйте,\n\n"
                "Подтверждаем доступность для встречи 10 марта 2026 года в нашем офисе в Санкт-Петербурге. "
                "Готовы принять в любое время после 11:30. Сообщите подтверждённое время, как только "
                "согласуете свой маршрут. Будем рады обсудить квартальные результаты и инвестиционные перспективы.\n\n"
                "С уважением,\nspb_partners@finance.ru"
            ),
        },
    ]
    with conn.cursor() as cur:
        for em in emails:
            cur.execute("""
                INSERT INTO email.messages
                    (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
                VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s)
            """, (
                folder_id,
                str(uuid.uuid4()),
                em["subject"], em["from_addr"], em["to_addr"], em["body"],
            ))
    conn.commit()
    print(f"[preprocess] Injected {len(emails)} emails.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        folder_id = get_or_create_inbox(conn)
        inject_emails(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
