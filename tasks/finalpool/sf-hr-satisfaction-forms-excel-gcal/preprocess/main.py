"""Preprocess script for sf-hr-satisfaction-forms-excel-gcal."""
import os
import argparse, json, os, sys, shutil, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM gcal.events")
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    # Шумовые события календаря (на русском). Ключевые слова для проверки
    # выживаемости шума синхронизированы с evaluation/main.py.
    cur.execute("""INSERT INTO gcal.events (summary, start_datetime, end_datetime, status)
        VALUES ('Командная планёрка', %s, %s, 'confirmed')""",
        (launch_dt.replace(hour=9), launch_dt.replace(hour=9, minute=30)))
    cur.execute("""INSERT INTO gcal.events (summary, start_datetime, end_datetime, status)
        VALUES ('Обеденный перерыв', %s, %s, 'confirmed')""",
        (launch_dt.replace(hour=12), launch_dt.replace(hour=13)))
    # Шумовая форма gform (id != опроса агента; типы вопросов валидны для CHECK-ограничения схемы).
    cur.execute("""INSERT INTO gform.forms (id, title, document_title, description)
        VALUES ('noise-form-001', 'Заявка на офисные принадлежности', 'Заявка на офисные принадлежности', 'Внутренняя форма заявки на офисные принадлежности')""")
    cur.execute("""INSERT INTO gform.questions (id, form_id, title, question_type, required, position)
        VALUES ('noise-q1', 'noise-form-001', 'Какие принадлежности вам нужны?', 'textQuestion', true, 0)""")
    cur.execute("""INSERT INTO gform.questions (id, form_id, title, question_type, required, position, config)
        VALUES ('noise-q2', 'noise-form-001', 'Уровень срочности?', 'choiceQuestion', false, 1, %s)""",
        (json.dumps({"type": "RADIO", "options": [{"value": "Низкий"}, {"value": "Высокий"}]}),))
    conn.commit()
    cur.close()
    conn.close()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)

if __name__ == "__main__":
    main()
