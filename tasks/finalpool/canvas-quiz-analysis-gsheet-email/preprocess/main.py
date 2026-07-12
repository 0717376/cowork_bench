"""
Скрипт предобработки для задачи canvas-quiz-analysis-gsheet-email.
Идемпотентно очищает данные Google Sheets и почты. Canvas — общая
read-only фикстура (источник данных), не модифицируется и не пересевается.
Ответы (таблица/документ/письмо) НЕ пред-засеваются — их создаёт агент.
"""
import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        # Идемпотентно очищаем данные Google Sheets
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.permissions")
        cur.execute("DELETE FROM gsheet.spreadsheets")
        print("[preprocess] Данные Google Sheets очищены.")

        # Идемпотентно очищаем данные почты
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Данные почты очищены.")

        conn.commit()
        print("[preprocess] Готово.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
