"""Preprocess для yf-sector-rotation-excel-notion-email (RU: moex + teamly).

Котировки moex.* засеяны глобально (db/zzz_moex_after_init.sql) и доступны только
для чтения — их не трогаем. Здесь только:
  - очищаем рабочие данные teamly.* (страницы агента, id > 3) и почту;
  - инжектим RU-шумовые страницы teamly и письма, чтобы агент учился отличать
    релевантное от постороннего.
Идемпотентно: повторный запуск не плодит дубликаты и не пре-сеет ответ.
"""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

# Шумовые страницы Teamly (RU): не относятся к секторной ротации.
NOISE_PAGES = [
    ("Протокол планёрки отдела", "Заметки с еженедельной планёрки: задачи, ответственные, сроки."),
    ("Маркетинговый план на 2026", "Стратегия и календарь маркетинговых активностей на год."),
    ("Бэклог спринта разработки", "Текущие задачи спринта и оценки трудозатрат."),
]

# Шумовые письма (RU).
NOISE_EMAILS = [
    ("admin@firm.com", ["ops@firm.com"], "Окно технического обслуживания",
     "Плановое обслуживание серверов в эти выходные с 02:00 до 06:00 МСК."),
    ("hr@firm.com", ["all_staff@firm.com"], "Напоминание о подаче заявлений на льготы",
     "Окно подачи заявлений закрывается в пятницу. Пожалуйста, отправьте свои выборы."),
]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        # --- Teamly: убираем страницы, созданные агентом/прошлым прогоном (id > 3) ---
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")

        # --- Почта: чистим, чтобы прошлые письма не засчитались ---
        cur.execute("SELECT to_regclass('email.attachments')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Очищены рабочие страницы teamly (id>3) и почта.")

        # --- Инжект RU-шумовых страниц Teamly в отдельное пространство ---
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('OPS', 'Операционные заметки', 'Разные рабочие заметки отдела.')
            ON CONFLICT (key) DO NOTHING
        """)
        cur.execute("SELECT id FROM teamly.spaces WHERE key='OPS'")
        ops_space = cur.fetchone()[0]
        for title, body in NOISE_PAGES:
            cur.execute("SELECT 1 FROM teamly.pages WHERE space_id=%s AND title=%s",
                        (ops_space, title))
            if cur.fetchone() is None:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s,%s,%s,%s)",
                    (ops_space, title, body, "admin"))

        # --- Инжект RU-шумовых писем ---
        cur.execute("SELECT id FROM email.folders WHERE name='Inbox' LIMIT 1")
        folder = cur.fetchone()
        if not folder:
            cur.execute("INSERT INTO email.folders (name) VALUES ('Inbox') RETURNING id")
            folder = cur.fetchone()
        folder_id = folder[0]

        import json
        for from_addr, to_addr, subject, body in NOISE_EMAILS:
            cur.execute("""
                INSERT INTO email.messages (folder_id, from_addr, to_addr, subject, body_text)
                VALUES (%s, %s, %s, %s, %s)
            """, (folder_id, from_addr, json.dumps(to_addr), subject, body))

        conn.commit()
        print("[preprocess] Инжектированы RU-шумовые страницы teamly и письма.")

    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
