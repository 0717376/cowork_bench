"""Preprocess для kulinar-event-registration-forms-gcal-teamly-email (RU-стек: kulinar/forms/teamly).

Готовит окружение и обеспечивает идемпотентность:
- Очищает gform.* (формы/вопросы/ответы) от прошлых прогонов.
- Очищает gcal.events и email.* от прошлых прогонов.
- Очищает пользовательские страницы teamly (id > сидовых) и пространство EVENTS.
- Создаёт пустое пространство teamly EVENTS, чтобы агенту было куда поместить
  страницу планирования мероприятия.

kulinar — read-only база рецептов, изменения не требуются.

ВАЖНО: НЕ создаёт заранее форму, событие, страницу или письмо, которые должен
произвести сам агент — это исключает авто-прохождение проверок.
"""

import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_writable_schemas(conn):
    """Очищает данные из изменяемых схем, используемых этой задачей."""
    cur = conn.cursor()

    # --- gform: формы/вопросы/ответы прошлых прогонов ---
    print("[preprocess] Очистка данных Forms (gform)...")
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")

    # --- gcal: события прошлых прогонов ---
    print("[preprocess] Очистка событий Google Calendar...")
    cur.execute("DELETE FROM gcal.events")

    # --- email: письма прошлых прогонов ---
    print("[preprocess] Очистка данных email...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")

    # --- teamly: убрать пользовательские страницы и пространство EVENTS ---
    # В zzz_teamly_after_init.sql засеяны сидовые пространства и страницы (id 1..3).
    print("[preprocess] Очистка пользовательских страниц teamly...")
    try:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    except Exception:
        pass
    try:
        cur.execute("DELETE FROM teamly.spaces WHERE key = 'EVENTS'")
    except Exception:
        pass

    # --- teamly: пустое пространство EVENTS для страницы планирования агента ---
    try:
        cur.execute(
            """
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('EVENTS', 'Мероприятия команды',
                    'Планы корпоративных мероприятий: меню, расписание, заметки.')
            ON CONFLICT (key) DO NOTHING
            """
        )
    except Exception:
        pass

    conn.commit()
    cur.close()
    print("[preprocess] Изменяемые схемы очищены; пространство teamly EVENTS обеспечено.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, help="Launch time")
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_writable_schemas(conn)
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        conn.close()

    print("[preprocess] Preprocessing completed successfully!")
