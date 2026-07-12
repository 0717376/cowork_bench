"""
Preprocess для rzd-insales-technoport-roadshow-spb-ekb-excel-teamly-email.

Что делает:
  - очищает email-таблицы, gcal.events
  - удаляет «свежие» страницы teamly (id > 3) — сиды трогать нельзя
  - наливает 1 событие в gcal (внутренний review накануне СПб-визита)
  - наливает 3 письма: запросы от СПб- и Екб-дистрибьюторов и пинг от руководителя

Требования:
  - PostgreSQL cowork_gym на хосте, доступ через PGHOST/PGDATABASE.
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

GCAL_EVENTS = [
    {
        "summary": "Внутренний product review перед road show",
        "description": (
            "Финальный внутренний обзор товаров и цен перед поездками в Санкт-Петербург "
            "и Екатеринбург на встречи с региональными дистрибьюторами."
        ),
        "start": "2026-03-09 14:00:00",
        "end": "2026-03-09 16:00:00",
    },
]

EMAILS = [
    {
        "message_id": "msg-spb-dist-001",
        "subject": "Доступность встречи — офис в Санкт-Петербурге",
        "from_addr": "spb_dist@partner.ru",
        "to_addr": ["bd@technoport.ru"],
        "date": "2026-03-05 09:30:00+00",
        "body_text": (
            "Здравствуйте,\n\n"
            "Получили ваш запрос про демонстрационный визит. Наш петербургский офис свободен "
            "10 марта 2026 года с 14:00 до 17:00. Подтвердите, пожалуйста, удобство времени. "
            "Особенно интересуют ваши линейки беспроводного аудио и мобильных аксессуаров. "
            "Пришлите, пожалуйста, заранее каталог товаров.\n\n"
            "С уважением,\nКоманда дистрибуции в СПб\nspb_dist@partner.ru"
        ),
    },
    {
        "message_id": "msg-ekb-dist-001",
        "subject": "Запрос каталога — Екатеринбург",
        "from_addr": "ekb_dist@partner.ru",
        "to_addr": ["bd@technoport.ru"],
        "date": "2026-03-06 11:00:00+00",
        "body_text": (
            "Уважаемая команда развития бизнеса,\n\n"
            "До нас дошёл слух, что «ТехноПорт» планирует road show по Уралу. Мы заинтересованы "
            "в статусе регионального дистрибьютора по топовым позициям. Пришлите, пожалуйста, "
            "актуальный товарный каталог и прайс. Особенно интересуют умные устройства и "
            "компьютерная периферия — на уральском рынке они хорошо идут.\n\n"
            "Ждём ответа,\nДистрибуция Екатеринбург\nekb_dist@partner.ru"
        ),
    },
    {
        "message_id": "msg-manager-001",
        "subject": "Запрос статуса по road show",
        "from_addr": "manager@company.ru",
        "to_addr": ["bd@technoport.ru"],
        "date": "2026-03-07 08:00:00+00",
        "body_text": (
            "Привет,\n\n"
            "Пришли, пожалуйста, апдейт по плану road show в Санкт-Петербурге и Екатеринбурге. "
            "Нужны маршрут поездок, список товаров для демонстрации и ключевые контакты. "
            "Подтверди также суммарные расходы на поездки и расписание встреч. "
            "Хочу всё пересмотреть до твоего отъезда.\n\n"
            "Спасибо,\nРуководитель\nmanager@company.ru"
        ),
    },
]


def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gcal.events")
        # teamly: чистим только пользовательские страницы, сиды (id<=3) сохраняем
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Очищены gcal, teamly.pages (id>3), email-таблицы.")


def inject_gcal_events(conn):
    with conn.cursor() as cur:
        for ev in GCAL_EVENTS:
            cur.execute("""
                INSERT INTO gcal.events (summary, description, start_datetime, end_datetime,
                    start_timezone, end_timezone, creator, organizer, attendees)
                VALUES (%s, %s, %s, %s, %s, %s, %s::jsonb, %s::jsonb, %s::jsonb)
            """, (
                ev["summary"], ev["description"], ev["start"], ev["end"],
                "Europe/Moscow", "Europe/Moscow",
                json.dumps({}), json.dumps({}), json.dumps([]),
            ))
    conn.commit()
    print(f"[preprocess] Залиты {len(GCAL_EVENTS)} событий в gcal.")


def inject_emails(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            conn.commit()
            cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
            row = cur.fetchone()
        folder_id = row[0]

        for em in EMAILS:
            cur.execute("""
                INSERT INTO email.messages (message_id, subject, from_addr, to_addr, date, body_text, folder_id)
                VALUES (%s, %s, %s, %s::jsonb, %s, %s, %s)
                ON CONFLICT (message_id) DO NOTHING
            """, (
                em["message_id"], em["subject"], em["from_addr"],
                json.dumps(em["to_addr"]), em["date"], em["body_text"], folder_id,
            ))
    conn.commit()
    print(f"[preprocess] Залиты {len(EMAILS)} писем.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_gcal_events(conn)
        inject_emails(conn)
    finally:
        conn.close()

    print("[preprocess] Preprocessing завершён успешно!")


if __name__ == "__main__":
    main()
