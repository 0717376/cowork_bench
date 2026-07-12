"""
Preprocess для yt-fireship-monthly-stats-excel-teamly (RU: notion -> teamly).

Очищает страницы Teamly с целевым заголовком, очищает таблицы почты,
гарантирует наличие папки INBOX и инжектит RU-шумовое пространство Teamly,
чтобы агент учился отличать релевантные данные от нерелевантных.

Данные YouTube доступны ТОЛЬКО ДЛЯ ЧТЕНИЯ — строки видео не трогаем,
чиним лишь метаданные channels.video_count (см. repair_channel_video_count).
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

TEAMLY_PAGE_TITLE = "Fireship Channel Analysis 2024-2025"

# RU-шумовые страницы Teamly: не относятся к анализу канала Fireship.
NOISE_PAGES = [
    ("Регламент учёта рабочего времени",
     "Памятка по заполнению табелей. К аналитике видеоканалов отношения не имеет."),
    ("Архив маркетинговых рассылок 2023",
     "Старые шаблоны email-кампаний. Не использовать в текущих отчётах."),
]


def clear_tables(conn):
    with conn.cursor() as cur:
        # Удаляем страницы Teamly агента/прошлого прогона с целевым заголовком.
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0]:
            cur.execute(
                "DELETE FROM teamly.pages WHERE title ILIKE %s",
                (f"%{TEAMLY_PAGE_TITLE}%",),
            )
        # Очищаем таблицы почты (sent_log до messages из-за FK).
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Очищены страницы teamly с целевым заголовком и таблицы почты.")


def repair_channel_video_count(conn):
    """Чинит youtube.channels.video_count по фактическому числу строк видео.

    Базовый сид содержит video_count=0 для всех каналов, из-за чего
    channels_listVideos считает totalPages=ceil(0/20)=0 и hasNextPage=false,
    и агенту доступна только первая страница (20 новейших видео). Метаданные
    восстанавливаются детерминированно из read-only строк видео.
    """
    with conn.cursor() as cur:
        cur.execute("""
            UPDATE youtube.channels c
            SET video_count = sub.cnt
            FROM (
                SELECT channel_id, COUNT(*) AS cnt
                FROM youtube.videos
                GROUP BY channel_id
            ) sub
            WHERE c.channel_id = sub.channel_id
        """)
        updated = cur.rowcount
    conn.commit()
    print(f"[preprocess] Починен youtube.channels.video_count для {updated} каналов.")


def seed_teamly_noise(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if not cur.fetchone()[0]:
            return
        # Шумовое пространство ANALYTICS_NOISE.
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('ANALYTICS_NOISE', 'Прочая аналитика', 'Шумовое пространство для проверки релевантности')
            ON CONFLICT (key) DO NOTHING
        """)
        cur.execute("SELECT id FROM teamly.spaces WHERE key='ANALYTICS_NOISE'")
        sid = cur.fetchone()[0]
        for title, body in NOISE_PAGES:
            cur.execute(
                "SELECT 1 FROM teamly.pages WHERE space_id=%s AND title=%s",
                (sid, title),
            )
            if not cur.fetchone():
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s,%s,%s,%s)",
                    (sid, title, body, "system"),
                )
    conn.commit()
    print("[preprocess] Инжектированы RU-шумовые страницы teamly.")


def ensure_email_folder(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            conn.commit()
            print("[preprocess] Создана папка почты INBOX.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        repair_channel_video_count(conn)
        seed_teamly_noise(conn)
        ensure_email_folder(conn)
    finally:
        conn.close()

    print("\n[preprocess] Препроцессинг успешно завершён!")


if __name__ == "__main__":
    main()
