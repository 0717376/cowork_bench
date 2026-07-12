"""Preprocess для insales-customer-gform-excel-notion (RU-стек: insales/forms/teamly).

- Очищает gform.* (формы/вопросы/ответы) от прошлых прогонов.
- Очищает пользовательские страницы teamly (id > сидовых).
- Инжектит шумовую форму (RU) и шумовую страницу teamly (RU) — источник «шума»,
  а не предзаполнение ответа. Реальные данные о клиентах агент читает живьём из wc.* (InSales).
"""
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
    # gform: убрать формы/вопросы/ответы прошлых прогонов
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    # teamly: убрать пользовательские страницы (засеяно 3 страницы с id 1..3)
    try:
        cur.execute("DELETE FROM teamly.page_labels WHERE page_id > 3")
    except Exception:
        pass
    try:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    except Exception:
        pass
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    # --- Шумовая форма gform (RU) ---
    cur.execute("""INSERT INTO gform.forms (id, title, document_title, description)
        VALUES ('noise-form-001', 'Заявка на канцелярию', 'Заявка на канцелярию', 'Внутренняя форма заявки на канцелярские товары')""")
    cur.execute("""INSERT INTO gform.questions (id, form_id, title, question_type, required, position)
        VALUES ('noise-q1', 'noise-form-001', 'Какие товары вам нужны?', 'TEXT', true, 0)""")
    cur.execute("""INSERT INTO gform.questions (id, form_id, title, question_type, required, position)
        VALUES ('noise-q2', 'noise-form-001', 'Уровень срочности?', 'RADIO', false, 1)""")
    # --- Шумовая страница teamly (RU) ---
    cur.execute("""INSERT INTO teamly.pages (space_id, title, body, author)
        VALUES (
            (SELECT id FROM teamly.spaces WHERE key='TEAM'),
            'Архив протоколов совещаний',
            E'# Архив протоколов совещаний\n\nЗдесь хранятся протоколы прошедших совещаний команды.',
            'admin'
        )""")
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
