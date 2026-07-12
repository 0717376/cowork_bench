"""Preprocess: подготовка Teamly перед выполнением задачи (clickhouse + teamly swap).

Идемпотентно: удаляем созданные пользователем страницы (seed имеет id <= 3) и
очищаем оставшиеся страницы 'Workforce Education Analysis', гарантируем наличие
пустого пространства под HR-аналитику. НЕ создаём целевую страницу заранее.
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = True
    cur = conn.cursor()

    print("[preprocess] Clearing Teamly data...")
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            # Удаляем пользовательские страницы (seed id <= 3) и любые ранее
            # созданные целевые страницы — идемпотентно, без pre-seed ответа.
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
            cur.execute(
                "DELETE FROM teamly.pages "
                "WHERE lower(title) LIKE %s OR lower(title) LIKE %s",
                ("%workforce education%", "%образовани%персонал%"),
            )
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('HREDU', 'Аналитика персонала',
                        'База знаний по аналитике образования и кадров.')
                ON CONFLICT (key) DO NOTHING
                """
            )
        print("[preprocess] Teamly data cleared.")
    except Exception as e:
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")

    cur.close()
    conn.close()
    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
