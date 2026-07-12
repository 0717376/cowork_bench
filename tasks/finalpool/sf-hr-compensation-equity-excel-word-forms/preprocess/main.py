"""
Preprocess script for sf-hr-compensation-equity-excel-word-gform task.
ClickHouse HR (sf_data schema) is read-only and russified centrally. Clear forms
and email data, inject noise forms.
"""
import argparse
import glob
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_gform(cur):
    print("[preprocess] Clearing Google Forms data...")
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    print("[preprocess] Google Forms data cleared.")


def clear_emails(cur):
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages WHERE folder_id != 0")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def inject_noise_forms(cur):
    print("[preprocess] Injecting noise forms...")
    cur.execute("""
        INSERT INTO gform.forms (id, title, description, created_at)
        VALUES
            ('noise_form_1', 'Заявка на канцелярию', 'Форма для заказа канцелярских принадлежностей', NOW() - INTERVAL '30 days'),
            ('noise_form_2', 'Регистрация на тимбилдинг', 'Запись на ежеквартальное мероприятие по тимбилдингу', NOW() - INTERVAL '15 days')
    """)
    cur.execute("""
        INSERT INTO gform.questions (id, form_id, title, question_type, config, position, required)
        VALUES
            ('nq1', 'noise_form_1', 'Какие принадлежности вам нужны?', 'textQuestion', '{}', 1, true),
            ('nq2', 'noise_form_1', 'Отдел', 'textQuestion', '{}', 2, true),
            ('nq3', 'noise_form_2', 'Вы придёте?', 'choiceQuestion', '{"type": "RADIO", "options": [{"value": "Да"}, {"value": "Нет"}, {"value": "Возможно"}]}', 1, true),
            ('nq4', 'noise_form_2', 'Ограничения в питании?', 'textQuestion', '{}', 2, false)
    """)
    print("[preprocess] Noise forms injected.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gform(cur)
        clear_emails(cur)
        inject_noise_forms(cur)
        conn.commit()
        print("[preprocess] DB cleanup and injection done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["Compensation_Equity.xlsx", "Equity_Report.docx"]:
            for f in glob.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
