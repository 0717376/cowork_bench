"""
Preprocess script for team-survey-report task.

Очищает изменяемые схемы и засевает:
1. Страницу Teamly «Engineering Team Projects» с 5 проектами (пространство ENG)
2. Форму «Team Satisfaction Survey» (схема gform) с 4 вопросами-оценками
3. 10 ответов на опрос по Leadership, Workload, Communication, Growth (1-5)

Идентификаторы (названия проектов, статусы, измерения, имена респондентов,
заголовок страницы и формы) намеренно оставлены на английском — их грепает
эвалюатор. Прозаические описания переведены на русский.
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


def clear_all(cur):
    """Очистка изменяемых схем в порядке, безопасном по FK."""
    print("[preprocess] Очистка изменяемых схем...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    # Teamly: удаляем только пользовательские записи (сид — 2 пространства, 3 страницы)
    try:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    except Exception:
        pass
    try:
        cur.execute("DELETE FROM teamly.spaces WHERE id > 2")
    except Exception:
        pass
    print("[preprocess] Схемы очищены.")


def inject_teamly_data(cur):
    """Создаёт пространство ENG и страницу «Engineering Team Projects» с 5 проектами."""
    print("[preprocess] Засев данных Teamly...")

    # (name, status, lead, deadline, описание на русском)
    projects = [
        ("Project Alpha", "Active", "Алиса Чен", "30 марта 2026",
         "Разработка системы аутентификации нового поколения с поддержкой OAuth 2.1 и passkey."),
        ("Project Beta", "Completed", "Борис Ванг", "28 февраля 2026",
         "Миграция легаси-монолита на микросервисную архитектуру. Успешно развёрнуто в продакшене."),
        ("Project Gamma", "Active", "Карина Ли", "15 апреля 2026",
         "Разработка дашборда аналитики в реальном времени для внутренних метрик и отслеживания KPI."),
        ("Project Delta", "On Hold", "Дмитрий Чжан", "1 мая 2026",
         "Редизайн мобильного приложения. Приостановлено до согласования направления дизайна со стейкхолдерами."),
        ("Project Epsilon", "Planning", "Ева Лю", "1 июня 2026",
         "Инициатива по оптимизации стоимости инфраструктуры. Сейчас на этапе сбора требований."),
    ]

    lines = [
        "# Engineering Team Projects",
        "",
        "На этой странице отслеживаются все текущие инженерные проекты: их статус, ответственные и сроки.",
        "",
    ]
    for name, status, lead, deadline, description in projects:
        lines.append(f"## {name}")
        lines.append(f"Status: {status}")
        lines.append(f"Lead: {lead}")
        lines.append(f"Deadline: {deadline}")
        lines.append(f"Описание: {description}")
        lines.append("")

    body = "\n".join(lines)

    cur.execute(
        """INSERT INTO teamly.spaces (key, name, description)
           VALUES ('ENG', 'Инженерная команда', 'Проекты, статусы и сроки инженерной команды.')
           RETURNING id""",
    )
    space_id = cur.fetchone()[0]
    cur.execute(
        """INSERT INTO teamly.pages (space_id, title, body, author)
           VALUES (%s, 'Engineering Team Projects', %s, 'people-ops')""",
        (space_id, body),
    )
    print(f"[preprocess] Создано пространство ENG (id={space_id}) со страницей по 5 проектам.")


def inject_form_and_responses(cur):
    """Создаёт форму опроса удовлетворённости и засевает 10 ответов."""
    print("[preprocess] Засев данных формы (схема gform)...")

    form_id = "form-team-satisfaction-001"
    q_name = "Your Name"
    q_leadership = "Leadership (1-5)"
    q_workload = "Workload (1-5)"
    q_communication = "Communication (1-5)"
    q_growth = "Growth (1-5)"

    cur.execute(
        """INSERT INTO gform.forms (id, title, document_title, description)
           VALUES (%s, %s, %s, %s)""",
        (
            form_id,
            "Team Satisfaction Survey",
            "Team Satisfaction Survey",
            "Пожалуйста, оцените следующие измерения по шкале от 1 до 5.",
        ),
    )

    rating_options = json.dumps({
        "type": "RADIO",
        "options": [
            {"value": "1"}, {"value": "2"}, {"value": "3"},
            {"value": "4"}, {"value": "5"},
        ],
    })

    questions = [
        (q_name, "textQuestion", True, "{}", 0),
        (q_leadership, "choiceQuestion", True, rating_options, 1),
        (q_workload, "choiceQuestion", True, rating_options, 2),
        (q_communication, "choiceQuestion", True, rating_options, 3),
        (q_growth, "choiceQuestion", True, rating_options, 4),
    ]

    for title, qtype, req, cfg, pos in questions:
        cur.execute(
            """INSERT INTO gform.questions
               (form_id, title, question_type, required, config, position)
               VALUES (%s, %s, %s, %s, %s::jsonb, %s)""",
            (form_id, title, qtype, req, cfg, pos),
        )

    # 10 ответов на опрос (имена-токены alice..jack оставлены английскими — их грепает эвал)
    # Alice: 4,3,5,4 | Bob: 3,4,3,3 | Carol: 5,2,4,5 | David: 3,3,3,2
    # Eva: 4,4,4,4 | Frank: 2,5,3,3 | Grace: 4,3,4,4 | Henry: 3,4,3,3
    # Irene: 4,3,4,4 | Jack: 3,4,3,3
    # Средние: Leadership=3.5, Workload=3.5, Communication=3.6, Growth=3.5
    responses = [
        ("Alice",  "4", "3", "5", "4"),
        ("Bob",    "3", "4", "3", "3"),
        ("Carol",  "5", "2", "4", "5"),
        ("David",  "3", "3", "3", "2"),
        ("Eva",    "4", "4", "4", "4"),
        ("Frank",  "2", "5", "3", "3"),
        ("Grace",  "4", "3", "4", "4"),
        ("Henry",  "3", "4", "3", "3"),
        ("Irene",  "4", "3", "4", "4"),
        ("Jack",   "3", "4", "3", "3"),
    ]

    for name, leadership, workload, communication, growth in responses:
        answers = json.dumps({
            q_name: name,
            q_leadership: leadership,
            q_workload: workload,
            q_communication: communication,
            q_growth: growth,
        })
        email = f"{name.lower()}@company.com"
        cur.execute(
            """INSERT INTO gform.responses (form_id, respondent_email, answers)
               VALUES (%s, %s, %s::jsonb)""",
            (form_id, email, answers),
        )

    print(f"[preprocess] Засеяна форма: {len(questions)} вопросов и {len(responses)} ответов.")


def ensure_email_folder(cur):
    """Гарантирует наличие папки INBOX."""
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if not row:
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_all(cur)
        inject_teamly_data(cur)
        inject_form_and_responses(cur)
        ensure_email_folder(cur)
        conn.commit()
        print("[preprocess] Операции с БД зафиксированы.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Ошибка БД: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("[preprocess] Препроцессинг успешно завершён!")


if __name__ == "__main__":
    main()
