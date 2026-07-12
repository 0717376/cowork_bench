"""Preprocess для terminal-clickhouse-teamly-forms-excel-email (RU-стек: clickhouse/teamly/forms).

Готовит окружение и обеспечивает идемпотентность:
- Очищает gform.* (формы/вопросы/ответы) от прошлых прогонов.
- Очищает email.* и инжектит RU-шумовые письма (не относятся к задаче).
- Очищает пользовательские страницы teamly (id > сидовых) и пространство HR.
- Создаёт пустое пространство teamly HR, куда агент поместит страницу-дашборд.
- Инжектит RU-шум: «лишняя» форма-опрос, «лишняя» страница teamly, шумовые письма.

ВАЖНО: НЕ создаёт заранее форму «Опрос вовлечённости сотрудников», страницу-дашборд,
Excel-файл или скрипт анализа, которые должен произвести сам агент — это исключает
авто-прохождение.

ClickHouse (sf_data) — read-only хранилище, сидится глобально и русифицируется
централизованным маппингом (db/zzz_clickhouse_after_init.sql). Здесь его не трогаем.
"""
import argparse
import glob
import json
import os
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_teamly(cur):
    print("[preprocess] Очистка пользовательских данных Teamly...")
    # В zzz_teamly_after_init.sql засеяно 2 пространства (TEAM, TRIPS) и 3 страницы (id 1..3).
    try:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    except Exception:
        pass
    try:
        cur.execute("DELETE FROM teamly.spaces WHERE key = 'HR'")
    except Exception:
        pass
    # Пустое пространство HR для страницы-дашборда (без содержимого ответа).
    try:
        cur.execute(
            """
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('HR', 'Отдел персонала',
                    'Аналитика вовлечённости, опросы и регламенты отдела персонала.')
            ON CONFLICT (key) DO NOTHING
            """
        )
    except Exception:
        pass
    print("[preprocess] Teamly: пространство HR обеспечено, старые страницы удалены.")


def clear_gform(cur):
    print("[preprocess] Очистка данных форм (gform)...")
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    print("[preprocess] Данные форм очищены.")


def clear_emails(cur):
    print("[preprocess] Очистка данных email...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages WHERE folder_id != 0")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] Данные email очищены.")


def inject_noise(cur):
    """Инжект RU-шума в email и teamly."""
    # Email-шум
    cur.execute("SELECT id FROM email.folders WHERE name='INBOX' LIMIT 1")
    row = cur.fetchone()
    if row:
        inbox_id = row[0]
    else:
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
        inbox_id = cur.fetchone()[0]
    noise_emails = [
        ("Еженедельное совещание персонала", "admin@company.com",
         json.dumps(["all@company.com"]), "Совещание завтра в 10:00."),
        ("Обновление по парковке", "facilities@company.com",
         json.dumps(["all@company.com"]), "Новые правила со следующего месяца."),
    ]
    for subj, from_addr, to_addr, body in noise_emails:
        cur.execute(
            "INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, body_text, is_read, date) "
            "VALUES (%s, %s, %s, %s, %s, %s, false, now())",
            (inbox_id, f"noise-{uuid.uuid4()}@company.com", subj, from_addr, to_addr, body))

    # Teamly-шум: лишняя страница в другом пространстве
    try:
        cur.execute(
            """
            INSERT INTO teamly.pages (space_id, title, body, author)
            VALUES ((SELECT id FROM teamly.spaces WHERE key='TEAM'),
                    'Протокол совещания (не по теме)',
                    E'# Протокол совещания\n\nОбсуждение канцелярских принадлежностей. К анализу вовлечённости отношения не имеет.',
                    'admin')
            """
        )
    except Exception:
        pass
    print("[preprocess] Инжектнут RU-шум (email + teamly).")


def inject_survey_responses(cur):
    """Инжект «лишней» (архивной) формы-опроса с 5 вопросами и 5 ответами.

    Это RU-шум: её заголовок НЕ совпадает с формой, которую создаёт агент
    («Опрос вовлечённости сотрудников»). Форма gform поддерживает только типы
    'choiceQuestion' и 'textQuestion' (нет нативного SCALE).
    """
    print("[preprocess] Инжект архивной формы-опроса (шум)...")
    form_id = str(uuid.uuid4())
    cur.execute(
        """
        INSERT INTO gform.forms (id, title, document_title, description)
        VALUES (%s, 'Архивный опрос (прошлый год)', 'Архивный опрос (прошлый год)',
                'Старые данные опроса для справки.')
        """,
        (form_id,),
    )

    q_ids = []
    # (заголовок, тип, config-json)
    questions = [
        ("Оценка удовлетворённости должностью", "choiceQuestion",
         '{"type":"RADIO","options":[{"value":"1"},{"value":"5"},{"value":"10"}]}'),
        ("Баланс работы и личной жизни", "choiceQuestion",
         '{"type":"RADIO","options":[{"value":"Доволен"},{"value":"Нейтрально"}]}'),
        ("Оценка карьерного роста", "choiceQuestion",
         '{"type":"RADIO","options":[{"value":"Хорошо"},{"value":"Плохо"}]}'),
        ("Оценка поддержки руководителя", "choiceQuestion",
         '{"type":"RADIO","options":[{"value":"1"},{"value":"5"}]}'),
        ("Рекомендация компании", "choiceQuestion",
         '{"type":"RADIO","options":[{"value":"Да"},{"value":"Нет"}]}'),
    ]
    for i, (title, qtype, config) in enumerate(questions):
        cur.execute(
            """
            INSERT INTO gform.questions (form_id, title, question_type, required, config, position)
            VALUES (%s, %s, %s, true, %s, %s) RETURNING id
            """,
            (form_id, title, qtype, config, i),
        )
        q_ids.append(cur.fetchone()[0])

    responses = [
        {"q0": "8", "q1": "Доволен", "q2": "Хорошо", "q3": "4", "q4": "Да"},
        {"q0": "6", "q1": "Нейтрально", "q2": "Удовлетворительно", "q3": "3", "q4": "Возможно"},
        {"q0": "9", "q1": "Очень доволен", "q2": "Отлично", "q3": "5", "q4": "Да"},
        {"q0": "5", "q1": "Недоволен", "q2": "Плохо", "q3": "2", "q4": "Нет"},
        {"q0": "7", "q1": "Доволен", "q2": "Хорошо", "q3": "4", "q4": "Да"},
    ]
    for j, resp in enumerate(responses):
        answers = {}
        for k, qid in enumerate(q_ids):
            answers[str(qid)] = {"questionId": str(qid), "textAnswers": {"answers": [{"value": resp[f"q{k}"]}]}}
        cur.execute(
            """
            INSERT INTO gform.responses (form_id, respondent_email, answers)
            VALUES (%s, %s, %s)
            """,
            (form_id, f"employee{j+1}@company.com", json.dumps(answers)),
        )

    print("[preprocess] Инжектнута архивная форма с 5 ответами (шум).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_teamly(cur)
        clear_gform(cur)
        clear_emails(cur)
        inject_survey_responses(cur)
        inject_noise(cur)
        conn.commit()
        print("[preprocess] Очистка БД завершена.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["Employee_Engagement_Report.xlsx", "engagement_analysis_output.txt"]:
            for f in glob.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Удалён {f}")

    print("[preprocess] Готово.")


if __name__ == "__main__":
    main()
