"""
Preprocess-скрипт для задачи canvas-exam-calendar-report.

Canvas доступен только на чтение, поэтому там ничего не меняем.
Этот скрипт:
1. Очищает почтовые данные (messages, attachments, sent_log, drafts)
2. Очищает события Google-календаря
3. Идемпотентно добавляет одно входящее письмо (на русском) с контекстом
   задачи для emily.watson — НЕ создаёт итоговое письмо, xlsx или события
   календаря (без пред-сидинга ответа).
"""

import os
import json
import argparse
from datetime import datetime, timedelta

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_emails(cur):
    """Очистить все почтовые данные, кроме структуры папок и конфигурации аккаунта."""
    print("[preprocess] Очистка почтовых данных...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Почтовые данные очищены.")


def clear_gcal(cur):
    """Очистить все события Google-календаря."""
    print("[preprocess] Очистка событий Google-календаря...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] События Google-календаря очищены.")


def seed_context_email(cur, launch_time):
    """
    Добавить одно входящее письмо на русском с контекстом задачи.
    Идемпотентно: messages только что очищены, поэтому дубликатов не будет;
    дополнительно вставляем по фиксированному message_id.
    Это НЕ итоговое письмо-сводка — это лишь напоминание от куратора.
    """
    try:
        launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        launch_dt = datetime.now()

    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    inbox_id = row[0] if row else 1

    body = (
        "Здравствуйте, Эмили!\n\n"
        "Напоминаю про подготовку к итоговым экзаменам осеннего семестра 2013 года. "
        "Пожалуйста, соберите план подготовки: по каждому курсу (код оканчивается на 2013J) "
        "найдите в Canvas задание Final Exam, занесите данные в таблицу, добавьте учебные "
        "сессии в календарь и пришлите сводку на адрес отдела.\n\n"
        "С уважением,\nКуратор учебного отдела"
    )

    # Идемпотентность: убираем возможную предыдущую копию по message_id.
    cur.execute(
        "DELETE FROM email.messages WHERE message_id = %s",
        ("<exam-context-001@openuniversity.ac.uk>",),
    )
    cur.execute(
        """INSERT INTO email.messages
           (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
           VALUES (%s, %s, %s, %s, %s, %s, %s, false)""",
        (
            inbox_id,
            "<exam-context-001@openuniversity.ac.uk>",
            "Напоминание: план подготовки к экзаменам (осень 2013)",
            "dept-admin@openuniversity.ac.uk",
            json.dumps(["emily.watson@openuniversity.ac.uk"]),
            launch_dt - timedelta(hours=2),
            body,
        ),
    )
    print("[preprocess] Контекстное письмо добавлено.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_emails(cur)
        clear_gcal(cur)
        seed_context_email(cur, args.launch_time)
        conn.commit()
        print("[preprocess] Готово.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Ошибка: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
