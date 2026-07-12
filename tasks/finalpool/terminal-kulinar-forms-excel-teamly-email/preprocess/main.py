"""Preprocess для terminal-kulinar-forms-excel-teamly-email (RU-стек: kulinar/forms/teamly).

Готовит окружение и обеспечивает идемпотентность:
- Очищает gform.* (формы/вопросы/ответы) от прошлых прогонов.
- Очищает email.* и инжектит RU-шумовое письмо (не относится к задаче).
- Очищает пользовательские страницы teamly (id > сидовых) и пространство WELLNESS.
- Создаёт пустое пространство teamly WELLNESS, куда агент поместит страницу базы знаний.
- Инжектит RU-шум: «лишняя» форма-опрос и «лишняя» страница teamly.

ВАЖНО: НЕ создаёт заранее форму-опрос, страницу базы знаний, Excel-файл или
menu_planner.py, которые должен произвести сам агент — это исключает авто-прохождение.
kulinar — read-only MCP (рецепты), сидится глобально; здесь ничего не трогаем.
"""
import argparse
import os
import uuid
import glob as globmod

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # --- gform: убрать формы/вопросы/ответы прошлых прогонов ---
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        print("[preprocess] Очищены данные gform.")

        # --- email: очистить и инжектнуть RU-шум ---
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Очищены данные email.")

        # --- teamly: убрать пользовательские страницы и пространство WELLNESS ---
        # В zzz_teamly_after_init.sql засеяно 2 пространства (TEAM, TRIPS) и 3 страницы (id 1..3).
        try:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        except Exception:
            pass
        try:
            cur.execute("DELETE FROM teamly.spaces WHERE key = 'WELLNESS'")
        except Exception:
            pass

        # --- teamly: пустое пространство WELLNESS для страницы базы знаний рецептов ---
        try:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('WELLNESS', 'Корпоративный велнес',
                        'Программа обедов для сотрудников: меню, рецепты, опросы.')
                ON CONFLICT (key) DO NOTHING
                """
            )
        except Exception:
            pass
        print("[preprocess] teamly: WELLNESS обеспечено, старые страницы удалены.")

        # --- Инжект RU-шума: лишняя форма-опрос ---
        noise_form_id = str(uuid.uuid4())
        cur.execute("""
            INSERT INTO gform.forms (id, title, document_title, description, created_at, updated_at)
            VALUES (%s, %s, %s, %s, NOW(), NOW())
        """, (noise_form_id, "Старый опрос сотрудников", "Старый опрос сотрудников",
              "Архивный опрос удовлетворённости сотрудников."))
        cur.execute("""
            INSERT INTO gform.questions (form_id, title, question_type, required, config, position)
            VALUES (%s, %s, %s, true, %s, 0)
        """, (noise_form_id, "Насколько вы довольны своим рабочим местом?",
              "choiceQuestion", '{"type":"RADIO","options":[{"value":"1"},{"value":"5"}]}'))
        print("[preprocess] Инжектнут RU-шум gform.")

        # --- Инжект RU-шума: лишняя страница teamly в другом пространстве ---
        try:
            cur.execute("""
                INSERT INTO teamly.pages (space_id, title, body, author)
                VALUES ((SELECT id FROM teamly.spaces WHERE key='TEAM'),
                        'Архив протоколов совещаний',
                        E'# Архив протоколов совещаний\n\nСтарые протоколы оперативок. К программе обедов отношения не имеют.',
                        'admin')
            """)
        except Exception:
            pass
        print("[preprocess] Инжектнут RU-шум teamly.")

        # --- Инжект RU-шумового письма ---
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            folder_id = row[0]
            cur.execute("SELECT COALESCE(MAX(id), 0) FROM email.messages")
            max_id = cur.fetchone()[0] + 1
            cur.execute("""
                INSERT INTO email.messages (id, folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                VALUES (%s, %s, %s, %s, %s, %s::jsonb, NOW(), %s, false)
            """, (max_id, folder_id, f"noise-{uuid.uuid4()}@company.com",
                  "Обновление по парковке", "facilities@company.com",
                  '["all_staff@company.com"]',
                  "Парковку перезаливают асфальтом на этих выходных. Пользуйтесь альтернативными местами."))
            print("[preprocess] Инжектнуто RU-шумовое письмо.")

        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # --- Идемпотентная очистка файлов-артефактов прошлого прогона ---
    if args.agent_workspace:
        for pattern in ["Meal_Program_Plan.xlsx", "menu_planner.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Удалён {f}")

    print("[preprocess] Готово.")


if __name__ == "__main__":
    main()
