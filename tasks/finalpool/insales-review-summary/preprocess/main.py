"""Preprocess: prepare Teamly workspace for a clean state (idempotent)."""
import os
import argparse
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Teamly (наш аналог Confluence) — целевое пространство и чистка прежних
    # результатов. НЕ создаём целевую страницу заранее (её создаёт агент).
    try:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('SHOP', 'Магазин',
                        'База знаний интернет-магазина: товары, отзывы, аналитика.')
                ON CONFLICT (key) DO NOTHING
                """
            )
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            # Идемпотентно удаляем прежнюю целевую страницу-дашборд, чтобы агент
            # создал её заново.
            cur.execute(
                "DELETE FROM teamly.pages WHERE title ILIKE %s",
                ("%product review dashboard%",),
            )
            # Шумовая страница-leftover, которую агент должен игнорировать и
            # которая НЕ должна удовлетворять проверке дашборда.
            cur.execute("SELECT id FROM teamly.spaces WHERE key = 'SHOP'")
            row = cur.fetchone()
            if row is None:
                cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
                row = cur.fetchone()
            space_id = row[0] if row else None
            if space_id is not None:
                cur.execute(
                    "SELECT 1 FROM teamly.pages WHERE title = %s",
                    ("Архив старых заметок по магазину",),
                )
                if cur.fetchone() is None:
                    cur.execute(
                        "INSERT INTO teamly.pages (space_id, title, body, author) "
                        "VALUES (%s, %s, %s, %s)",
                        (space_id, "Архив старых заметок по магазину",
                         "Устаревшие черновики. Не относится к текущей задаче.",
                         "shop"),
                    )
    except Exception as e:
        print(f"[preprocess] WARNING: teamly setup skipped: {e}")

    conn.commit()
    cur.close()
    conn.close()
    print("Teamly workspace prepared (space SHOP ensured, dashboard page cleared).")


if __name__ == "__main__":
    main()
