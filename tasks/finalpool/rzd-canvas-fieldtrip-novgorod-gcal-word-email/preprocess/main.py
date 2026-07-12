"""Preprocess for rzd-canvas-fieldtrip-novgorod-gcal-word-email.

Очищает gcal/email и canvas.announcements курса 9991, инжектит контекст:
- 1 событие в gcal: "Обычное занятие — Изучение культурного наследия" 2026-03-12 09:00-11:00
- 1 входящее письмо от students@university.ru с просьбой прислать уведомление
Canvas-курс «Изучение культурного наследия России» (id=9991) и его 12
enrollments создаются один раз в db/zzz_rzd_after_init.sql; preprocess их не трогает.
"""
import argparse
import json
import os
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

COURSE_ID = 9991


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
        # Только объявления тестового курса — чужие не трогаем.
        cur.execute("DELETE FROM canvas.announcements WHERE course_id = %s", (COURSE_ID,))
    conn.commit()
    print("[preprocess] Очищены gcal, email, canvas.announcements (курс 9991).")


def ensure_email_folder(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        fid = cur.fetchone()[0]
    conn.commit()
    return fid


def inject_gcal(conn):
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO gcal.events (summary, description, start_datetime, end_datetime,
                start_timezone, end_timezone)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            "Обычное занятие — Изучение культурного наследия",
            "Регулярная утренняя лекция курса по культурному наследию России. Сессия 12 марта.",
            "2026-03-12 09:00:00",
            "2026-03-12 11:00:00",
            "Europe/Moscow",
            "Europe/Moscow",
        ))
    conn.commit()
    print("[preprocess] Создано событие в GCal: Обычное занятие 12.03.")


def inject_email(conn, folder_id):
    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO email.messages
                (folder_id, message_id, subject, from_addr, to_addr, date, body_text)
            VALUES (%s, %s, %s, %s, %s::jsonb, %s, %s)
        """, (
            folder_id,
            "msg-students-novgorod-001",
            "Вопрос по экскурсии в Великий Новгород — Культурное наследие",
            "students@university.ru",
            json.dumps(["professor@university.ru"]),
            "2026-03-06 14:00:00+03",
            (
                "Здравствуйте,\n\nМы слышали, что Вы планируете для нашего курса по "
                "культурному наследию учебную экскурсию в Великий Новгород — посмотреть "
                "Софийский собор, Новгородский кремль и Юрьев монастырь. Это будет очень "
                "ценный опыт. Не могли бы Вы прислать официальное уведомление об экскурсии "
                "со всеми деталями: расписанием поездов, стоимостью на студента и тем, что "
                "нужно взять с собой? Хотелось бы получить как можно скорее, чтобы успеть "
                "подготовиться.\n\nСпасибо,\nСтуденты курса по культурному наследию"
            ),
        ))
    conn.commit()
    print("[preprocess] Создано входящее письмо от students@university.ru.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_gcal(conn)
        folder_id = ensure_email_folder(conn)
        inject_email(conn, folder_id)
    finally:
        conn.close()

    print("[preprocess] Подготовка для rzd-canvas-fieldtrip-novgorod-gcal-word-email завершена.")


if __name__ == "__main__":
    main()
