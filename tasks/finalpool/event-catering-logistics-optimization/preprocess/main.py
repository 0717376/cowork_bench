"""Preprocess для event-catering-logistics-optimization (RU-стек: kulinar/forms).

Готовит окружение идемпотентно:
- Очищает gform.* (формы/вопросы/ответы) от прошлых прогонов.
- Очищает email.* и gcal.events, чтобы проверки видели только то, что создал агент.
- Инжектит ОДНО входящее письмо от events@company.ru с просьбой подготовить план
  кейтеринга (контекст мероприятия), не раскрывая итоговых блюд/чисел.

ВАЖНО: НЕ создаёт заранее форму, Excel-файл, Word-документ, итоговое письмо или события
календаря, которые должен произвести сам агент — это исключает авто-прохождение проверок.
Исходные данные (participant_list.xlsx / menu_templates.xlsx / Event_Brief.md) лежат в
initial_workspace и копируются агенту средствами харнеса.
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
        # gform: убрать любые формы/вопросы/ответы прошлых прогонов
        for tbl in ("gform.responses", "gform.questions", "gform.forms"):
            try:
                cur.execute(f"DELETE FROM {tbl}")
            except Exception:
                conn.rollback()
        # email / gcal
        for tbl in ("email.attachments", "email.sent_log", "email.messages",
                    "email.drafts", "gcal.events"):
            try:
                cur.execute(f"DELETE FROM {tbl}")
            except Exception:
                conn.rollback()
    conn.commit()
    print("[preprocess] Очищены gform.*, email.*, gcal.events.")


def get_or_create_inbox(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        conn.commit()
        return cur.fetchone()[0]


def inject_email(conn, folder_id):
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
            VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s)
            """,
            (
                folder_id,
                str(uuid.uuid4()),
                "Конференция «Импульс-2026» — нужен план кейтеринга",
                "events@company.ru",
                json.dumps(["coordinator@company.ru"]),
                (
                    "Здравствуйте,\n\n"
                    "Мы проводим ежегодную конференцию «Импульс-2026» в Москве на 500 человек. "
                    "Просим подготовить полный план кейтеринга и логистики.\n\n"
                    "Что нужно сделать:\n"
                    "1. Подобрать 6–8 блюд из базы рецептов с учётом всех пищевых ограничений участников.\n"
                    "2. Запустить опрос пищевых предпочтений участников.\n"
                    "3. Свести план в Excel (меню, сводка по диетам, бюджет).\n"
                    "4. Подготовить предложение для руководства в Word с резервным планом.\n"
                    "5. Согласовать дедлайны с поставщиками через календарь и прислать итог ответом на это письмо.\n\n"
                    "Бюджет питания: 1500 ₽ на человека (итого 750 000 ₽).\n\n"
                    "Спасибо,\nОтдел мероприятий\nevents@company.ru"
                ),
            ),
        )
    conn.commit()
    print("[preprocess] Создано входящее письмо от events@company.ru.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        os.makedirs(args.agent_workspace, exist_ok=True)

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        folder_id = get_or_create_inbox(conn)
        inject_email(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Подготовка event-catering-logistics-optimization завершена.")


if __name__ == "__main__":
    main()
