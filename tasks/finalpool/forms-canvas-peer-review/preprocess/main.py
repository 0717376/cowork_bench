"""
Preprocess для gform-canvas-peer-review (RU-стек: forms/teamly + canvas/gsheet/excel).

Canvas — read-only, изменений там нет.
Этот скрипт:
1. Очищает пользовательские схемы (gform, gsheet) и пользовательские страницы teamly.
2. Создаёт форму "Group Project Peer Review" с 6 вопросами.
3. Инжектит 20 ответов взаимного оценивания для 6 студентов (исходные данные, не ответ агента).
4. Обеспечивает пустое пространство teamly REVIEWS для итоговой страницы агента.

ВАЖНО: НЕ создаёт заранее Excel-файл, Google-таблицу или страницу Teamly,
которые должен произвести сам агент — это исключает авто-прохождение проверок.
Имена студентов остаются на английском, чтобы совпадать с ростером Canvas (read-only).
"""

import os
import argparse
import json

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_writable_schemas(cur):
    """Очистить пользовательские схемы в FK-безопасном порядке."""
    print("[preprocess] Очистка пользовательских схем...")
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.permissions")
    cur.execute("DELETE FROM gsheet.sheets")
    cur.execute("DELETE FROM gsheet.spreadsheets")
    # teamly: удалить пользовательские страницы (сидовые id 1..3) и пространство REVIEWS
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
            cur.execute("DELETE FROM teamly.spaces WHERE key = 'REVIEWS'")
    except Exception as e:
        print(f"[preprocess] teamly cleanup skipped: {e}")
    print("[preprocess] Пользовательские схемы очищены.")


def ensure_teamly_space(cur):
    """Создать пустое пространство teamly для итоговой страницы агента."""
    try:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is None:
            return
        cur.execute(
            """INSERT INTO teamly.spaces (key, name, description)
               VALUES ('REVIEWS', 'Взаимное оценивание',
                       'Итоги взаимного оценивания студентов по групповым проектам.')
               ON CONFLICT (key) DO NOTHING"""
        )
        print("[preprocess] Пространство teamly REVIEWS обеспечено.")
    except Exception as e:
        print(f"[preprocess] teamly space skipped: {e}")


def inject_form_and_responses(cur):
    """Создать форму взаимного оценивания и инжектить 20 ответов."""
    print("[preprocess] Инжект формы и ответов взаимного оценивания...")

    form_id = "form-peer-review-001"
    q_reviewer = "q-reviewer-001"
    q_reviewee = "q-reviewee-001"
    q_contribution = "q-contribution-001"
    q_communication = "q-communication-001"
    q_quality = "q-quality-001"
    q_comments = "q-comments-001"

    # Create form (название/описание формы — на английском: substring-проверки eval)
    cur.execute(
        """INSERT INTO gform.forms (id, title, document_title, description)
           VALUES (%s, %s, %s, %s)""",
        (
            form_id,
            "Group Project Peer Review",
            "Group Project Peer Review",
            "Пожалуйста, оцените участников вашей группы по вкладу (Contribution), "
            "коммуникации (Communication) и качеству работы (Quality of Work) "
            "в групповом проекте по курсу Biochemistry & Bioinformatics.",
        ),
    )

    # Create questions (заголовки вопросов — на английском, eval ищет по ним)
    questions = [
        (q_reviewer, form_id, "Your Name", "textQuestion", True, "{}", 0),
        (q_reviewee, form_id, "Person Being Reviewed", "textQuestion", True, "{}", 1),
        (
            q_contribution,
            form_id,
            "Contribution Score (1-5)",
            "choiceQuestion",
            True,
            json.dumps({
                "type": "RADIO",
                "options": [
                    {"value": "1"},
                    {"value": "2"},
                    {"value": "3"},
                    {"value": "4"},
                    {"value": "5"},
                ],
            }),
            2,
        ),
        (
            q_communication,
            form_id,
            "Communication Score (1-5)",
            "choiceQuestion",
            True,
            json.dumps({
                "type": "RADIO",
                "options": [
                    {"value": "1"},
                    {"value": "2"},
                    {"value": "3"},
                    {"value": "4"},
                    {"value": "5"},
                ],
            }),
            3,
        ),
        (
            q_quality,
            form_id,
            "Quality of Work (1-5)",
            "choiceQuestion",
            True,
            json.dumps({
                "type": "RADIO",
                "options": [
                    {"value": "1"},
                    {"value": "2"},
                    {"value": "3"},
                    {"value": "4"},
                    {"value": "5"},
                ],
            }),
            4,
        ),
        (q_comments, form_id, "Comments", "textQuestion", False, "{}", 5),
    ]

    for q in questions:
        cur.execute(
            """INSERT INTO gform.questions
               (id, form_id, title, question_type, required, config, position)
               VALUES (%s, %s, %s, %s, %s, %s::jsonb, %s)""",
            q,
        )

    # 20 ответов взаимного оценивания для 6 студентов.
    # Имена студентов — на английском (совпадение с ростером Canvas, read-only).
    # Тексты комментариев — на русском (RU-проза, источник данных, не ответ агента).
    # Alice Wong: высокий результат (~4.5), Frank Liu: низкий (~2.3, flagged),
    # остальные — умеренные (3.0 - 4.2).
    responses = [
        # Alice Wong оценивает: Bob, Carol, David
        ("Alice Wong", "Bob Martinez", "4", "3", "4", "Боб надёжно выполнил лабораторную часть."),
        ("Alice Wong", "Carol Zhang", "4", "4", "5", "Кэрол отлично справилась с отчётом."),
        ("Alice Wong", "David Kim", "3", "4", "3", "Дэвиду стоило бы больше вложиться в код."),

        # Bob Martinez оценивает: Alice, Eva, Frank
        ("Bob Martinez", "Alice Wong", "5", "5", "4", "Элис эффективно руководила группой."),
        ("Bob Martinez", "Eva Patel", "4", "3", "4", "Анализ данных у Евы был тщательным."),
        ("Bob Martinez", "Frank Liu", "2", "2", "3", "Фрэнк пропустил несколько встреч."),

        # Carol Zhang оценивает: Alice, David, Frank
        ("Carol Zhang", "Alice Wong", "5", "4", "5", "Элис была движущей силой нашего проекта."),
        ("Carol Zhang", "David Kim", "4", "3", "4", "Дэвид прилично поработал над презентацией."),
        ("Carol Zhang", "Frank Liu", "2", "3", "2", "Фрэнк сдал свою часть с опозданием, потребовались правки."),

        # David Kim оценивает: Alice, Bob, Frank, Eva
        ("David Kim", "Alice Wong", "4", "5", "5", "Превосходное лидерство со стороны Элис."),
        ("David Kim", "Bob Martinez", "3", "4", "3", "Боб был надёжен, но недостаточно инициативен."),
        ("David Kim", "Frank Liu", "3", "2", "2", "Качество работы Фрэнка было ниже ожиданий."),
        ("David Kim", "Eva Patel", "3", "4", "3", "Ева хорошо общалась, но могла бы повысить отдачу."),

        # Eva Patel оценивает: Alice, Bob, Carol, Frank
        ("Eva Patel", "Alice Wong", "5", "4", "4", "Элис всегда держала нас в графике."),
        ("Eva Patel", "Bob Martinez", "4", "4", "4", "Боб был надёжным членом команды."),
        ("Eva Patel", "Carol Zhang", "3", "4", "4", "Текст у Кэрол получился сильным."),
        ("Eva Patel", "Frank Liu", "2", "2", "2", "Фрэнк не тянул свою часть нагрузки."),

        # Frank Liu оценивает: Alice, Carol, David
        ("Frank Liu", "Alice Wong", "4", "5", "5", "Элис была великолепна."),
        ("Frank Liu", "Carol Zhang", "4", "3", "4", "Кэрол помогла с оформлением текста."),
        ("Frank Liu", "David Kim", "3", "3", "4", "Дэвид нормально справился со своими частями."),
    ]

    for reviewer, reviewee, contrib, comm, quality, comments in responses:
        answers = json.dumps({
            q_reviewer: reviewer,
            q_reviewee: reviewee,
            q_contribution: contrib,
            q_communication: comm,
            q_quality: quality,
            q_comments: comments,
        })
        # respondent_email — стабильный @-формат на основе имени оценивающего
        email = reviewer.lower().replace(" ", ".") + "@university.edu"
        cur.execute(
            """INSERT INTO gform.responses (form_id, respondent_email, answers)
               VALUES (%s, %s, %s::jsonb)""",
            (form_id, email, answers),
        )

    print(f"[preprocess] Инжектировано: форма + {len(questions)} вопросов и {len(responses)} ответов.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_writable_schemas(cur)
        ensure_teamly_space(cur)
        inject_form_and_responses(cur)
        conn.commit()
        print("[preprocess] Изменения в БД зафиксированы.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Ошибка БД: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("[preprocess] Предобработка успешно завершена!")


if __name__ == "__main__":
    main()
