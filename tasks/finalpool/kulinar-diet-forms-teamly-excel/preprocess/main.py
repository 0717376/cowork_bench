"""Preprocess для kulinar-diet-forms-teamly-excel (RU-стек: kulinar/forms/teamly).

Готовит окружение и обеспечивает идемпотентность:
- Очищает gform.* (формы/вопросы/ответы) от прошлых прогонов.
- Очищает пользовательские страницы teamly (id > сидовых) и пространство RECIPES.
- Очищает email.messages.
- Создаёт пустое пространство teamly RECIPES, чтобы агенту было куда поместить
  страницу базы знаний по рецептам.

ВАЖНО: НЕ создаёт заранее форму, страницу базы знаний, Excel-файл или письмо,
которые должен произвести сам агент — это исключает авто-прохождение проверок.
Рецепты kulinar засеяны глобально, отдельный db-файл не нужен.
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
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        # --- gform: убрать любые формы/вопросы/ответы прошлых прогонов ---
        try:
            cur.execute("DELETE FROM gform.responses")
        except Exception:
            conn.rollback()
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")

        # --- email: очистить отправленные письма ---
        try:
            cur.execute("DELETE FROM email.messages")
        except Exception:
            conn.rollback()

        # --- teamly: убрать пользовательские страницы и пространство RECIPES ---
        # В zzz_teamly_after_init.sql засеяно 2 пространства (TEAM, TRIPS) и 3 страницы.
        try:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        except Exception:
            conn.rollback()
        try:
            cur.execute("DELETE FROM teamly.spaces WHERE key = 'RECIPES'")
        except Exception:
            conn.rollback()

        # --- teamly: пустое пространство RECIPES под страницу базы знаний агента ---
        try:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('RECIPES', 'База знаний по рецептам',
                        'Полезные рецепты для программы здорового питания команды.')
                ON CONFLICT (key) DO NOTHING
                """
            )
        except Exception:
            conn.rollback()

        conn.commit()
        print("[preprocess] gform/email очищены; teamly: RECIPES обеспечено, старые страницы удалены.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
