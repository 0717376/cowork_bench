"""
Preprocess для задачи canvas-submission-timeline-gcal.

Canvas доступен только для чтения (живой сервер, server-side seed). Этот скрипт
лишь идемпотентно очищает события общего календаря (gcal.events). Ответы НЕ
пред-засеиваются: документ и события календаря создаёт агент.
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


def clear_gcal(cur):
    print("[preprocess] Очистка событий Google Calendar...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] События Google Calendar очищены.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gcal(cur)
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
