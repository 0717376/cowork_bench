"""Preprocess for course-enrollment-analytics-dashboard.

Что делает (идемпотентно, БЕЗ преднасева ответов):
  - Гарантирует наличие пространства Teamly "Учебный отдел" (ключ REGISTRAR),
    в котором агент создаст страницу "At-Risk Intervention Plan"; удаляет
    прежние страницы плана/набора, оставшиеся от предыдущих прогонов.
  - Очищает писательские схемы: gcal.events, email.*, аналитические таблицы
    google sheet с названием про набор/риск.
  - Засевает RU-«шум»: пара служебных писем во входящих и пара событий в
    календаре, которые агент НЕ должен трогать (reverse-noise проверки).

Canvas засеян глобально (общий набор курсов/зачислений/успеваемости) — мы его
НЕ трогаем: и агент, и оценка читают эти данные «вживую». Никакие ответы
(страница плана, лист метрик, письмо, событие) здесь НЕ создаются.
"""
import os
import argparse
import json
from datetime import datetime, timedelta

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")


def setup_teamly(cur):
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; skipping teamly setup.")
        return
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('REGISTRAR', 'Учебный отдел',
                'Аналитика набора студентов и планы вмешательства для группы риска.')
        ON CONFLICT (key) DO NOTHING
    """)
    # Идемпотентность: удаляем прежние страницы плана/набора.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%at-risk intervention plan%'
            OR title ILIKE '%intervention plan%'
            OR title ILIKE '%план%вмешат%'
    """)
    print("[preprocess] Teamly ready: 'REGISTRAR' space ensured, prior plan pages cleared.")


def clear_gsheet(cur):
    cur.execute("SELECT to_regclass('gsheet.spreadsheets')")
    if cur.fetchone()[0] is None:
        return
    cur.execute("""
        SELECT id FROM gsheet.spreadsheets
         WHERE title ILIKE '%course enrollment analytics%'
            OR title ILIKE '%enrollment analytics%'
            OR title ILIKE '%набор%аналит%'
    """)
    ids = [r[0] for r in cur.fetchall()]
    for sid in ids:
        cur.execute("DELETE FROM gsheet.cells WHERE spreadsheet_id = %s", (sid,))
        cur.execute("DELETE FROM gsheet.sheets WHERE spreadsheet_id = %s", (sid,))
        cur.execute("DELETE FROM gsheet.permissions WHERE spreadsheet_id = %s", (sid,))
        cur.execute("DELETE FROM gsheet.spreadsheets WHERE id = %s", (sid,))
    if ids:
        print(f"[preprocess] Cleared {len(ids)} prior analytics spreadsheet(s).")


def clear_writable(cur):
    cur.execute("DELETE FROM gcal.events")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")


def inject_noise(cur, launch_time):
    try:
        launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError):
        launch_dt = datetime(2026, 6, 1, 10, 0, 0)

    # Шумовые события календаря — агент НЕ должен их удалять.
    cur.execute("""INSERT INTO gcal.events (summary, start_datetime, end_datetime, status)
        VALUES ('Планёрка', %s, %s, 'confirmed')""",
        (launch_dt.replace(hour=9, minute=0), launch_dt.replace(hour=9, minute=30)))
    cur.execute("""INSERT INTO gcal.events (summary, start_datetime, end_datetime, status)
        VALUES ('Обед', %s, %s, 'confirmed')""",
        (launch_dt.replace(hour=12, minute=0), launch_dt.replace(hour=13, minute=0)))

    # Шумовые письма во входящих — агент НЕ должен их пересылать/отправлять.
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    inbox_id = row[0] if row else 1
    cur.execute("""INSERT INTO email.messages
        (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
        VALUES (%s, '<noise-001@univ.edu>', 'Еженедельная рассылка деканата',
                'newsletter@university.edu', %s, %s,
                'Новости и анонсы учебного отдела за неделю.', true)""",
        (inbox_id, json.dumps(['all@university.edu']), launch_dt - timedelta(hours=5)))
    cur.execute("""INSERT INTO email.messages
        (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
        VALUES (%s, '<noise-002@univ.edu>', 'Плановое обслуживание LMS',
                'it@university.edu', %s, %s,
                'В субботу запланированы технические работы в системе обучения.', false)""",
        (inbox_id, json.dumps(['staff@university.edu']), launch_dt - timedelta(hours=3)))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-06-01 10:00:00")
    args = parser.parse_args()

    if args.agent_workspace:
        from pathlib import Path
        Path(args.agent_workspace).mkdir(parents=True, exist_ok=True)

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        setup_teamly(cur)
        clear_gsheet(cur)
        clear_writable(cur)
        inject_noise(cur, args.launch_time)
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
