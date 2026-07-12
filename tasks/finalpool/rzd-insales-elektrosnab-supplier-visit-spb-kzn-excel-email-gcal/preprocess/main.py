"""
Preprocess для rzd-insales-elektrosnab-supplier-visit-spb-kzn-excel-email-gcal.

Что делает:
  - очищает email-таблицы и gcal.events
  - наливает 1 событие в gcal (внутренний брифинг команды закупок перед поездкой в СПб)
  - наливает 3 письма-инициатора: запросы от СПб- и Казань-поставщиков и пинг от руководителя

Требования:
  - PostgreSQL cowork_gym на хосте, доступ через PGHOST/PGDATABASE.
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
        cur.execute("DELETE FROM gcal.events")
    conn.commit()
    print("[preprocess] Очищены email-таблицы и gcal.events.")


def get_or_create_inbox(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        conn.commit()
        return cur.fetchone()[0]


def inject_gcal_event(conn):
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO gcal.events
                (id, summary, start_datetime, end_datetime, start_timezone, end_timezone,
                 creator, organizer, attendees, description, location)
            VALUES (%s, %s, %s::timestamptz, %s::timestamptz, %s, %s,
                    %s::jsonb, %s::jsonb, %s::jsonb, %s, %s)
        """, (
            str(uuid.uuid4()),
            "Внутренний брифинг команды закупок перед визитами к поставщикам",
            "2026-03-09T15:00:00+03:00",
            "2026-03-09T16:00:00+03:00",
            "Europe/Moscow",
            "Europe/Moscow",
            json.dumps({"email": "procurement@elektrosnab.ru"}),
            json.dumps({"email": "procurement@elektrosnab.ru"}),
            json.dumps([{"email": "buyer@elektrosnab.ru"}]),
            "Финальный брифинг по целям визитов в Санкт-Петербург и Казань. Сверка списка поставщиков и обсуждаемых позиций.",
            "Москва, головной офис ЭлектроСнаб, переговорная Б",
        ))
    conn.commit()
    print("[preprocess] Залит 1 gcal event (брифинг).")


def inject_emails(conn, folder_id):
    emails = [
        {
            "message_id": "msg-spb-supplier-001",
            "subject": "Re: Визит закупки — даты в марте",
            "from_addr": "spb_supplier@partner.ru",
            "to_addr": json.dumps(["buyer@elektrosnab.ru"]),
            "body": (
                "Здравствуйте,\n\n"
                "Спасибо, что рассматриваете нас как поставщика. По вашим последним заказам "
                "видно стабильный рост — готовы обсудить долгосрочный контракт. Наша команда в "
                "Санкт-Петербурге свободна 10 марта 2026 с 11:00 до 18:00 в офисе на Васильевском "
                "острове. Подтвердите, пожалуйста, удобное время после прибытия «Сапсана» и "
                "пришлите заранее список интересующих позиций.\n\n"
                "С уважением,\nКоманда поставок СПб\nspb_supplier@partner.ru"
            ),
        },
        {
            "message_id": "msg-kzn-supplier-001",
            "subject": "Запрос каталога — Казань",
            "from_addr": "kzn_supplier@partner.ru",
            "to_addr": json.dumps(["buyer@elektrosnab.ru"]),
            "body": (
                "Уважаемая служба закупок,\n\n"
                "Наш казанский склад готов принять вас 17 марта 2026 после полудня. Пришлите, "
                "пожалуйста, актуальный список позиций, по которым у вас критичные остатки — "
                "подготовим коммерческое предложение и спецификации к встрече. Особенно "
                "интересуют категории беспроводного аудио и мобильных аксессуаров.\n\n"
                "Ждём подтверждения,\nКазанский поставщик\nkzn_supplier@partner.ru"
            ),
        },
        {
            "message_id": "msg-procurement-head-001",
            "subject": "Запрос статуса по визитам к поставщикам",
            "from_addr": "procurement@elektrosnab.ru",
            "to_addr": json.dumps(["buyer@elektrosnab.ru"]),
            "body": (
                "Привет,\n\n"
                "Пришли, пожалуйста, апдейт по плану визитов в Санкт-Петербург и Казань. "
                "Нужны: список из пяти приоритетных позиций с поставщиками, маршрут поездок "
                "(поезда и тарифы Бизнес-класса), расписание встреч с учётом 30-минутного "
                "буфера после прибытия. Сводку оформи в Excel-файле Supplier_Visit_Plan.xlsx "
                "и пришли итоговое письмо на этот адрес.\n\n"
                "Спасибо,\nРуководитель закупок\nprocurement@elektrosnab.ru"
            ),
        },
    ]
    with conn.cursor() as cur:
        for em in emails:
            cur.execute("""
                INSERT INTO email.messages
                    (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
                VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s)
                ON CONFLICT (message_id) DO NOTHING
            """, (
                folder_id,
                em["message_id"],
                em["subject"], em["from_addr"], em["to_addr"], em["body"],
            ))
    conn.commit()
    print(f"[preprocess] Залиты {len(emails)} писем.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_gcal_event(conn)
        folder_id = get_or_create_inbox(conn)
        inject_emails(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Preprocessing завершён успешно!")


if __name__ == "__main__":
    main()
