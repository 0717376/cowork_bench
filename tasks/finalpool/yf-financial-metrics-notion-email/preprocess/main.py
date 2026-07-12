"""Preprocess для yf-financial-metrics-notion-email (RU: moex + teamly).

Котировки и финотчётность moex.* засеяны глобально и доступны только для чтения —
их не трогаем. Здесь только:
  - очищаем рабочие страницы Teamly (id > 3), созданные агентом/прошлым прогоном;
  - очищаем почту.
Идемпотентно: повторный запуск не плодит дубликаты и не пре-сеет ответ.
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


def clear_teamly(cur):
    print("[preprocess] Очистка рабочих страниц Teamly (id>3)...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    print("[preprocess] Страницы Teamly очищены.")


def clear_emails(cur):
    print("[preprocess] Очистка почты...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Почта очищена.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_teamly(cur)
        clear_emails(cur)
        conn.commit()
        print("[preprocess] Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
