"""Preprocess для portfolio-diversification-analysis-engine (RU / moex-finance).

- Данные moex.* засеяны глобально (db/zzz_moex_after_init.sql), доступны только
  для чтения — инъекция/очистка не требуется.
- Копируем исходный holdings.csv в рабочую директорию агента (источник, который
  агент потребляет; это НЕ файл-ответ).
- Идемпотентно чистим writable-схемы email.* и gcal.* и подсыпаем нейтральный
  RU-шум, чтобы агент находил нужные письмо/событие сам. Файлы-ответы
  (.docx/.pptx) НЕ создаём.
"""
import argparse
import json
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)


def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        cur.execute("DELETE FROM gcal.events")
        conn.commit()
    finally:
        cur.close()
        conn.close()


def inject_noise(launch_time):
    """Нейтральный RU-шум: одно письмо и одно событие, не относящиеся к задаче."""
    conn = get_conn()
    cur = conn.cursor()
    try:
        launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")

        inbox_id = 1
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            inbox_id = row[0]
        cur.execute(
            """INSERT INTO email.messages
               (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
               VALUES (%s, %s, %s, %s, %s, %s, %s, true)""",
            (inbox_id, '<noise-pda-001@capital-invest.ru>', 'Еженедельный дайджест',
             'newsletter@capital-invest.ru', json.dumps(['all@capital-invest.ru']),
             launch_dt - timedelta(hours=5),
             'Внутренние новости компании за неделю. К задаче отношения не имеет.'),
        )

        cur.execute(
            """INSERT INTO gcal.events (summary, start_datetime, end_datetime, description, status)
               VALUES ('Ежедневная планёрка', %s, %s, 'Регулярная планёрка отдела', 'confirmed')""",
            (launch_dt.replace(hour=9, minute=0), launch_dt.replace(hour=9, minute=15)),
        )
        conn.commit()
    finally:
        cur.close()
        conn.close()


def seed_workspace(agent_workspace):
    if not agent_workspace:
        return
    agent_ws = Path(agent_workspace)
    agent_ws.mkdir(parents=True, exist_ok=True)
    src = os.path.join(TASK_ROOT, "initial_workspace", "holdings.csv")
    if os.path.exists(src):
        shutil.copy(src, agent_ws / "holdings.csv")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-05-30 10:00:00")
    args = parser.parse_args()

    seed_workspace(args.agent_workspace)
    try:
        clear_writable_schemas()
        inject_noise(args.launch_time)
    except Exception as e:
        print(f"[preprocess] предупреждение: подготовка БД пропущена: {e}")
    print("[preprocess] готово: holdings.csv засеян, email/gcal очищены, шум добавлен")


if __name__ == "__main__":
    main()
