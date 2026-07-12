"""Preprocess for terminal-canvas-gform-gcal-excel-email.
Clears gform, gcal, and email data. Injects noise emails and a GForm with responses.
"""
import argparse
import json
import uuid

import os
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "cowork_gym"), user="eigent", password="camel")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        # Clear writable schemas
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        cur.execute("DELETE FROM email.drafts")
        conn.commit()
        print("[preprocess] Cleared gform, gcal, email schemas.")

        # Inject noise emails
        cur.execute("SELECT id FROM email.folders WHERE name='INBOX' LIMIT 1")
        inbox_id = cur.fetchone()[0]

        noise_emails = [
            ("Напоминание о собрании персонала", "admin@university.edu",
             json.dumps(["all-staff@university.edu"]),
             "Напоминание: еженедельное собрание персонала завтра в 10:00 в аудитории 204."),
            ("Техническое обслуживание библиотечной системы", "it@university.edu",
             json.dumps(["faculty@university.edu"]),
             "В эту субботу библиотечная система будет недоступна из-за технического обслуживания."),
            ("Обновление правил парковки на кампусе", "facilities@university.edu",
             json.dumps(["all@university.edu"]),
             "Новые правила парковки вступают в силу в следующем месяце. Пожалуйста, ознакомьтесь с обновлённой политикой."),
            ("Срок подачи заявок на исследовательский грант", "grants@university.edu",
             json.dumps(["researchers@university.edu"]),
             "Крайний срок подачи заявок на исследовательские гранты весны 2026 — 15 апреля."),
        ]
        for subj, from_addr, to_addr, body in noise_emails:
            cur.execute(
                "INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, "
                "body_text, is_read, date) VALUES (%s, %s, %s, %s, %s, %s, false, now())",
                (inbox_id, f"noise-{uuid.uuid4()}@university.edu", subj, from_addr, to_addr, body)
            )
        conn.commit()
        print("[preprocess] Injected 4 noise emails.")

        # Inject a noise GForm with study habit responses (existing survey, a distractor)
        form_id = str(uuid.uuid4())
        cur.execute(
            "INSERT INTO gform.forms (id, title, document_title) VALUES (%s, %s, %s)",
            (form_id, "Опрос о привычках подготовки к экзаменам прошлого семестра",
             "Опрос о привычках подготовки")
        )
        q1_id = str(uuid.uuid4())
        q2_id = str(uuid.uuid4())
        q3_id = str(uuid.uuid4())
        cur.execute(
            "INSERT INTO gform.questions (id, form_id, title, question_type, position) "
            "VALUES (%s, %s, %s, %s, %s)",
            (q1_id, form_id, "Как вы обычно готовитесь к тестам?", "textQuestion", 1)
        )
        cur.execute(
            "INSERT INTO gform.questions (id, form_id, title, question_type, position) "
            "VALUES (%s, %s, %s, %s, %s)",
            (q2_id, form_id, "Оцените свой уровень уверенности (1-5)", "choiceQuestion", 2)
        )
        cur.execute(
            "INSERT INTO gform.questions (id, form_id, title, question_type, position) "
            "VALUES (%s, %s, %s, %s, %s)",
            (q3_id, form_id, "Предпочитаемое время для учёбы", "choiceQuestion", 3)
        )
        # Add some responses
        for i in range(5):
            resp_id = str(uuid.uuid4())
            answers = {
                q1_id: f"Ответ {i+1}: повторение конспектов и решение задач",
                q2_id: str(3 + (i % 3)),
                q3_id: ["Утро", "Вечер", "День", "Ночь", "Выходные"][i]
            }
            cur.execute(
                "INSERT INTO gform.responses (id, form_id, respondent_email, answers) "
                "VALUES (%s, %s, %s, %s)",
                (resp_id, form_id, f"student{i+1}@university.edu", json.dumps(answers))
            )
        conn.commit()
        print("[preprocess] Injected noise GForm with 5 responses.")

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
