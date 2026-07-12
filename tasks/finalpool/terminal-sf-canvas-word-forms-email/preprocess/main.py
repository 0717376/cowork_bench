"""Preprocess for terminal-sf-canvas-word-gform-email.
Clears gform and email. Injects noise emails and a noise form.
Does NOT pre-seed the "Training Feedback Survey" form: the agent must create it
(its existence/questions/options are the agent's deliverable, checked by eval).
The survey response data the analysis consumes ships as a static workspace input
file (initial_workspace/survey_responses.json), not as a pre-built gform form."""
import argparse
import json
import os
import uuid

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")


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
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            conn.rollback()
            # Re-clear in case rollback undid previous deletes
            cur.execute("DELETE FROM gform.responses")
            cur.execute("DELETE FROM gform.questions")
            cur.execute("DELETE FROM gform.forms")
            cur.execute("DELETE FROM email.attachments")
            cur.execute("DELETE FROM email.sent_log")
            cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared gform, email schemas.")

        # Inject noise emails
        cur.execute("SELECT id FROM email.folders WHERE name='INBOX' LIMIT 1")
        inbox_id = cur.fetchone()[0]

        noise_emails = [
            ("Расписание праздничных дней", "admin@company.com",
             json.dumps(["all@company.com"]),
             "Просьба ознакомиться с обновлённым графиком выходных на II квартал 2026 года."),
            ("Окно технического обслуживания ИТ-систем", "it@company.com",
             json.dumps(["all@company.com"]),
             "Плановое обслуживание в субботу с 2:00 до 6:00."),
            ("Квартальный отчёт о выручке", "finance@company.com",
             json.dumps(["leadership@company.com"]),
             "Выручка за IV квартал превысила план на 12%. Полный отчёт во вложении."),
            ("Адаптация новых сотрудников", "hr@company.com",
             json.dumps(["managers@company.com"]),
             "Встречаем 15 новых сотрудников, которые выходят в следующий понедельник. Просьба подготовить рабочие места."),
            ("Ремонт парковки", "facilities@company.com",
             json.dumps(["all@company.com"]),
             "Парковка B будет закрыта на ремонт с 15 по 17 марта."),
        ]
        for subj, from_addr, to_addr, body in noise_emails:
            cur.execute(
                "INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, "
                "body_text, is_read, date) VALUES (%s, %s, %s, %s, %s, %s, false, now())",
                (inbox_id, f"noise-{uuid.uuid4()}@company.com", subj, from_addr, to_addr, body)
            )
        conn.commit()
        print("[preprocess] Injected 5 noise emails.")

        # NOTE: The "Training Feedback Survey" form is the agent's deliverable and is
        # intentionally NOT pre-seeded here. The agent must create it (5 questions with
        # the specified types/options) using the forms MCP. The survey response data the
        # analysis consumes is provided as a static workspace input file
        # (survey_responses.json), shipped via initial_workspace/.

        # Inject a noise form
        noise_form_id = str(uuid.uuid4())
        cur.execute(
            "INSERT INTO gform.forms (id, title, document_title) VALUES (%s, %s, %s)",
            (noise_form_id, "Опрос о снеках в офисе", "Опрос о снеках")
        )
        nq1 = str(uuid.uuid4())
        cur.execute(
            "INSERT INTO gform.questions (id, form_id, title, question_type, position) "
            "VALUES (%s, %s, %s, %s, %s)",
            (nq1, noise_form_id, "Любимый вид снеков?", "TEXT", 1)
        )
        for i in range(3):
            cur.execute(
                "INSERT INTO gform.responses (id, form_id, respondent_email, answers) "
                "VALUES (%s, %s, %s, %s)",
                (str(uuid.uuid4()), noise_form_id, f"snacker{i+1}@company.com",
                 json.dumps({nq1: ["Чипсы", "Фрукты", "Печенье"][i]}))
            )
        conn.commit()
        print("[preprocess] Injected noise form.")

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
