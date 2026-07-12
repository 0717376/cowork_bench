"""Preprocess для kulinar-catering-forms-teamly-excel (RU-стек: kulinar/forms/teamly).

Готовит окружение и обеспечивает идемпотентность:
- Очищает gform.* (формы/вопросы/ответы), оставшиеся от предыдущих прогонов.
- Очищает пользовательские страницы teamly (id > сидовых) и пространство EVENTS.
- Создаёт пустое пространство teamly EVENTS, чтобы агенту было куда поместить страницу меню.

ВАЖНО: НЕ создаёт заранее форму, страницу меню или Excel-файл, которые должен
произвести сам агент — это исключает авто-прохождение проверок.
"""
import os
import argparse
import psycopg2

DB = dict(
    host=os.environ.get("PGHOST", "localhost"),
    port=5432,
    dbname=os.environ.get("PGDATABASE", "cowork_gym"),
    user="eigent",
    password="camel",
)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        # --- gform: убрать любые формы/вопросы/ответы прошлых прогонов ---
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")

        # --- teamly: убрать пользовательские страницы и пространство EVENTS ---
        # В zzz_teamly_after_init.sql засеяно 2 пространства (TEAM, TRIPS) и 3 страницы.
        # Чистим всё, что мог насоздавать агент в прошлый раз.
        try:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        except Exception:
            pass
        try:
            cur.execute("DELETE FROM teamly.spaces WHERE key = 'EVENTS'")
        except Exception:
            pass

        # --- teamly: пустое пространство EVENTS для страницы меню агента ---
        try:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('EVENTS', 'Мероприятия команды',
                        'Планы корпоративных мероприятий: меню, бюджеты, заметки.')
                ON CONFLICT (key) DO NOTHING
                """
            )
        except Exception:
            pass

        conn.commit()
        print("[preprocess] gform очищен; teamly: EVENTS обеспечено, старые страницы удалены.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
