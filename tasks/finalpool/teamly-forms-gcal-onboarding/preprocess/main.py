"""
Preprocess для задачи notion-gform-gcal-onboarding (RU: teamly + forms).

Очищает пользовательские схемы и засевает исходные данные, которые потребляет агент:
1. Страница teamly "New Employee Onboarding Checklist" с 6 пунктами чек-листа
   (в пространстве KB). Заголовок страницы остаётся английским — его грепает eval/задача.
2. Форма gform "New Employee Information" с 5 вопросами.
3. 3 ответа формы (новые сотрудники: Анна Парк, Михаил Чен, Мария Родригес;
   почты @company.com сохранены, отделы — на русском).

ВАЖНО: preprocess НЕ создаёт результат, который должен произвести агент
(презентацию, события, письма, раздел "March 2026 New Hires").
"""

import os
import argparse
import json

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# Заголовок остаётся английским — его грепают задача и eval.
CHECKLIST_PAGE_TITLE = "New Employee Onboarding Checklist"

# Пункты чек-листа на русском (прозу агент перенесёт в слайды).
CHECKLIST_ITEMS = [
    "Оформить документы HR",
    "Настроить рабочее место и учётные записи",
    "Познакомиться с командой",
    "Посетить вводный семинар",
    "Изучить корпоративный справочник",
    "Пройти комплаенс-обучение",
]

FORM_TITLE = "New Employee Information"

# Новые сотрудники: (ФИО, email@company.com, отдел_ru, дата_выхода, экстренный_контакт)
NEW_HIRES = [
    ("Анна Парк", "sarah.park@company.com", "Инженерия", "2026-03-16", "Иван Парк: 555-0101"),
    ("Михаил Чен", "mike.chen@company.com", "Продажи", "2026-03-16", "Лариса Чен: 555-0202"),
    ("Мария Родригес", "amy.rodriguez@company.com", "Маркетинг", "2026-03-16", "Карл Родригес: 555-0303"),
]


def clear_all(cur):
    """Очистка пользовательских схем в FK-безопасном порядке (идемпотентно)."""
    print("[preprocess] Очистка пользовательских схем...")
    for stmt in (
        "DELETE FROM email.attachments",
        "DELETE FROM email.sent_log",
        "DELETE FROM email.drafts",
        "DELETE FROM email.messages",
        "DELETE FROM gform.responses",
        "DELETE FROM gform.questions",
        "DELETE FROM gform.forms",
        "DELETE FROM gcal.events",
        # Teamly: чистим только пользовательские страницы/пространства
        # (в zzz_teamly_after_init.sql сидовых пространств 2, страниц 3).
        "DELETE FROM teamly.pages WHERE id > 3",
        "DELETE FROM teamly.spaces WHERE id > 2",
    ):
        try:
            cur.execute(stmt)
        except Exception as e:
            print(f"[preprocess]   пропуск '{stmt}': {e}")
    print("[preprocess] Схемы очищены.")


def inject_teamly_checklist(cur):
    """Создаёт пространство KB и страницу с чек-листом онбординга."""
    print("[preprocess] Засев Teamly...")

    cur.execute(
        """INSERT INTO teamly.spaces (key, name, description)
           VALUES ('KB', 'База знаний', 'Внутренние регламенты и чек-листы компании.')
           RETURNING id"""
    )
    space_id = cur.fetchone()[0]

    body_lines = [f"# {CHECKLIST_PAGE_TITLE}", "",
                  "Шаги, которые должен пройти каждый новый сотрудник:", ""]
    for item in CHECKLIST_ITEMS:
        body_lines.append(f"- {item}")
    body = "\n".join(body_lines)

    cur.execute(
        """INSERT INTO teamly.pages (space_id, title, body, author)
           VALUES (%s, %s, %s, 'hr-admin')""",
        (space_id, CHECKLIST_PAGE_TITLE, body),
    )
    print(f"[preprocess] Создана страница «{CHECKLIST_PAGE_TITLE}» "
          f"в пространстве KB (id={space_id}) с {len(CHECKLIST_ITEMS)} пунктами.")


def inject_form_data(cur):
    """Создаёт форму регистрации и 3 ответа новых сотрудников."""
    print("[preprocess] Засев Google-формы (gform)...")

    form_id = "form-new-employee-001"

    cur.execute(
        """INSERT INTO gform.forms (id, title, document_title, description)
           VALUES (%s, %s, %s, %s)""",
        (
            form_id,
            FORM_TITLE,
            FORM_TITLE,
            "Пожалуйста, заполните свои данные до даты выхода на работу.",
        ),
    )

    # Названия вопросов остаются английскими (Full Name/Email/Department/...) —
    # это ключи answers, которые могут грепаться.
    questions = [
        ("Full Name", "textQuestion", True, "{}", 0),
        ("Email", "textQuestion", True, "{}", 1),
        (
            "Department",
            "choiceQuestion",
            True,
            json.dumps(
                {
                    "type": "RADIO",
                    "options": [
                        {"value": "Инженерия"},
                        {"value": "Продажи"},
                        {"value": "Маркетинг"},
                        {"value": "HR"},
                    ],
                },
                ensure_ascii=False,
            ),
            2,
        ),
        ("Start Date", "textQuestion", True, "{}", 3),
        ("Emergency Contact", "textQuestion", True, "{}", 4),
    ]

    for title, qtype, req, cfg, pos in questions:
        cur.execute(
            """INSERT INTO gform.questions
               (form_id, title, question_type, required, config, position)
               VALUES (%s, %s, %s, %s, %s::jsonb, %s)""",
            (form_id, title, qtype, req, cfg, pos),
        )

    for name, email, dept, start_date, emergency in NEW_HIRES:
        answers = json.dumps(
            {
                "Full Name": name,
                "Email": email,
                "Department": dept,
                "Start Date": start_date,
                "Emergency Contact": emergency,
            },
            ensure_ascii=False,
        )
        cur.execute(
            """INSERT INTO gform.responses (form_id, respondent_email, answers)
               VALUES (%s, %s, %s::jsonb)""",
            (form_id, email, answers),
        )

    print(f"[preprocess] Засеяна форма «{FORM_TITLE}»: "
          f"{len(questions)} вопросов и {len(NEW_HIRES)} ответов.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_all(cur)
        inject_teamly_checklist(cur)
        inject_form_data(cur)
        conn.commit()
        print("[preprocess] Изменения в БД зафиксированы.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Ошибка БД: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("[preprocess] Подготовка завершена успешно!")


if __name__ == "__main__":
    main()
