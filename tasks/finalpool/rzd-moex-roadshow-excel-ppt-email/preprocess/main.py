"""
Preprocess для задачи roadshow (РЖД + moex-finance).
Идемпотентно очищает таблицы почты. Данные РЖД (схема rzd.*) и moex-finance
(схема moex.*) засеяны глобально (zzz_rzd_after_init.sql, zzz_moex_after_init.sql)
и доступны только для чтения — здесь их не трогаем и новых db-файлов не добавляем.
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
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Cleared email tables.")


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

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
