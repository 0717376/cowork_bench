"""
Препроцессинг для задачи team-retreat (rzd + teamly + word + gcal + email).
Идемпотентно очищает gcal, email и созданные агентом страницы Teamly.
Данные о поездах в схеме rzd.* — READ-ONLY (засеяны глобально).
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


def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gcal.events")
        try:
            cur.execute("DELETE FROM email.attachments")
        except Exception:
            pass
        try:
            cur.execute("DELETE FROM email.sent_log")
        except Exception:
            pass
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
        # Сидовые страницы teamly (3 шт., id 1..3 из db/zzz_teamly_after_init.sql)
        # оставляем; удаляем только то, что мог создать агент (id > 3).
        try:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Очищены gcal.events, email и страницы teamly, созданные агентом.")


def ensure_email_folder(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            conn.commit()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        ensure_email_folder(conn)
    finally:
        conn.close()

    print("[preprocess] Подготовка завершена успешно!")


if __name__ == "__main__":
    main()
