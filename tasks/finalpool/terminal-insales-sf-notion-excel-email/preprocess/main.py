"""Preprocess for terminal-insales-sf-notion-excel-email (russified).
Clears teamly user-pages and email schemas idempotently, injects RU noise data.
Does NOT pre-create the 'Support Quality Tracker' space or any audit output."""
import argparse
import json
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}

# Noise space/pages must NOT satisfy the audit deliverable checks.
NOISE_SPACE_KEY = "PROJ_NOISE"
NOISE_SPACE_NAME = "Трекер проектов"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        # Clear teamly user-created pages (keep seed pages id <= 3); ensure schema present.
        print("[preprocess] Clearing teamly user pages...")
        cur.execute("SELECT to_regclass('teamly.pages')")
        teamly_pages = cur.fetchone()[0] is not None
        cur.execute("SELECT to_regclass('teamly.spaces')")
        teamly_spaces = cur.fetchone()[0] is not None
        if teamly_pages:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        if teamly_spaces:
            # Remove a previously injected noise space (cascade drops its pages).
            cur.execute("DELETE FROM teamly.spaces WHERE key = %s", (NOISE_SPACE_KEY,))

        # Clear email
        print("[preprocess] Clearing email schema...")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")

        # Inject RU noise teamly space + pages (leftovers the agent must ignore).
        if teamly_spaces and teamly_pages:
            cur.execute(
                "INSERT INTO teamly.spaces (key, name, description) "
                "VALUES (%s, %s, %s) ON CONFLICT (key) DO NOTHING RETURNING id",
                (NOISE_SPACE_KEY, NOISE_SPACE_NAME,
                 "Пространство для отслеживания текущих проектов отдела."))
            row = cur.fetchone()
            if row is None:
                cur.execute("SELECT id FROM teamly.spaces WHERE key = %s",
                            (NOISE_SPACE_KEY,))
                row = cur.fetchone()
            noise_space_id = row[0]

            for title, owner in [
                ("Редизайн сайта", "Алиса"),
                ("Мобильное приложение v2", "Борис"),
                ("Обновление конвейера данных", "Виктор"),
            ]:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) "
                    "VALUES (%s, %s, %s, %s)",
                    (noise_space_id, title,
                     f"Статус: В работе. Ответственный: {owner}. "
                     "Старые заметки по проекту, к текущей задаче не относятся.",
                     owner))

        # Inject RU noise emails
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            inbox_id = row[0]
        else:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            inbox_id = cur.fetchone()[0]

        cur.execute("SELECT id FROM email.folders WHERE name = 'Sent' LIMIT 1")
        row = cur.fetchone()
        if row:
            sent_id = row[0]
        else:
            cur.execute("INSERT INTO email.folders (name) VALUES ('Sent') RETURNING id")
            sent_id = cur.fetchone()[0]

        noise_emails = [
            ("Заметки с еженедельной планёрки", "manager@company.com", "team@company.com",
             "Заметки с планёрки за эту неделю. Просьба ознакомиться до понедельника."),
            ("Утверждение бюджета на Q1", "finance@company.com", "vp_cx@company.com",
             "Бюджет на Q1 утверждён. Можно приступать к запланированным инициативам."),
            ("График адаптации новых сотрудников", "hr@company.com", "team@company.com",
             "Приветствуем новых коллег. Сессии адаптации начинаются в понедельник."),
            ("Окно технического обслуживания сервера", "ops@company.com", "all@company.com",
             "Плановые работы в субботу с 2:00 до 6:00. Возможны кратковременные простои."),
            ("Сводка отзывов клиентов — январь", "analytics@company.com", "product_team@company.com",
             "Январская сводка отзывов во вложении. Общий настрой положительный."),
        ]

        for subj, from_addr, to_addr, body in noise_emails:
            import uuid
            msg_id = f"<noise-{uuid.uuid4()}@company.com>"
            cur.execute("""
                INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                VALUES (%s, %s, %s, %s, %s, NOW(), %s, true)
            """, (inbox_id, msg_id, subj, from_addr,
                  json.dumps([to_addr]), body))

        conn.commit()
        print("[preprocess] Done. RU noise injected into teamly and email.")

    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
