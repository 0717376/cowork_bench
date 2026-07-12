"""
Preprocess для задачи kulinar-survey-analysis (RU-стек: forms/kulinar).

Очищает изменяемые схемы и инжектит:
1. Google-форму «Опрос предпочтений для командного обеда» с 5 вопросами (схема gform.*)
2. 12 ответов с разными вкусовыми предпочтениями.

Категории блюд соответствуют категориям базы рецептов kulinar
(горячее/суп/гарнир/салат/десерт), чтобы агент мог найти реальные рецепты
самой популярной категории через getRecipesByCategory.
"""

import argparse
import json
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_all(cur):
    """Clear all writable schemas in FK-safe order."""
    print("[preprocess] Clearing all writable schemas...")
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
    cur.execute("DELETE FROM gcal.events")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.sheets")
    cur.execute("DELETE FROM gsheet.spreadsheets")
    print("[preprocess] All schemas cleared.")


def inject_form_data(cur):
    """Inject the Team Lunch Preference Survey form and 12 responses."""
    print("[preprocess] Injecting form data...")

    form_id = "form-lunch-survey-001"
    q_name = "q-name-001"
    q_cuisine = "q-cuisine-001"
    q_spice = "q-spice-001"
    q_allergy = "q-allergy-001"
    q_difficulty = "q-difficulty-001"

    # Create form
    cur.execute(
        """INSERT INTO gform.forms (id, title, document_title, description)
           VALUES (%s, %s, %s, %s)""",
        (form_id, "Опрос предпочтений для командного обеда",
         "Опрос предпочтений для командного обеда",
         "Помогите спланировать следующий командный обед — поделитесь своими вкусовыми предпочтениями."),
    )

    # Create questions.
    # Варианты «типа блюда» соответствуют категориям базы kulinar
    # (горячее/суп/гарнир/салат/десерт).
    questions = [
        (q_name, form_id, "Ваше имя", "textQuestion", True, "{}", 0),
        (
            q_cuisine, form_id, "Предпочитаемый тип блюда", "choiceQuestion", True,
            json.dumps({
                "type": "RADIO",
                "options": [
                    {"value": "горячее"},
                    {"value": "суп"},
                    {"value": "гарнир"},
                    {"value": "салат"},
                    {"value": "десерт"},
                ],
            }),
            1,
        ),
        (
            q_spice, form_id, "Переносимость остроты", "choiceQuestion", True,
            json.dumps({
                "type": "RADIO",
                "options": [
                    {"value": "Слабо"},
                    {"value": "Средне"},
                    {"value": "Остро"},
                ],
            }),
            2,
        ),
        (q_allergy, form_id, "Есть ли пищевые аллергии?", "textQuestion", False, "{}", 3),
        (
            q_difficulty, form_id, "Максимальная сложность готовки (1-5)", "choiceQuestion", True,
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
    ]

    for q in questions:
        cur.execute(
            """INSERT INTO gform.questions
               (id, form_id, title, question_type, required, config, position)
               VALUES (%s, %s, %s, %s, %s, %s::jsonb, %s)""",
            q,
        )

    # 12 responses
    # горячее=6, суп=2, гарнир=2, салат=1, десерт=1
    responses = [
        ("Анна", "anna@company.com", "горячее", "Средне", "Нет", "3"),
        ("Борис", "boris@company.com", "суп", "Слабо", "Арахис", "2"),
        ("Виктория", "viktoria@company.com", "горячее", "Остро", "Нет", "4"),
        ("Дмитрий", "dmitry@company.com", "гарнир", "Средне", "Нет", "2"),
        ("Елена", "elena@company.com", "горячее", "Средне", "Морепродукты", "3"),
        ("Фёдор", "fedor@company.com", "салат", "Слабо", "Нет", "2"),
        ("Галина", "galina@company.com", "суп", "Слабо", "Нет", "1"),
        ("Игорь", "igor@company.com", "горячее", "Остро", "Нет", "5"),
        ("Ирина", "irina@company.com", "десерт", "Слабо", "Молочное", "2"),
        ("Кирилл", "kirill@company.com", "горячее", "Средне", "Нет", "3"),
        ("Ксения", "ksenia@company.com", "гарнир", "Средне", "Глютен", "3"),
        ("Леонид", "leonid@company.com", "горячее", "Остро", "Нет", "4"),
    ]

    for name, email, cuisine, spice, allergy, difficulty in responses:
        answers = json.dumps({
            q_name: name,
            q_cuisine: cuisine,
            q_spice: spice,
            q_allergy: allergy,
            q_difficulty: difficulty,
        })
        cur.execute(
            """INSERT INTO gform.responses (form_id, respondent_email, answers)
               VALUES (%s, %s, %s::jsonb)""",
            (form_id, email, answers),
        )

    print(f"[preprocess] Injected form with {len(questions)} questions and {len(responses)} responses.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_all(cur)
        inject_form_data(cur)
        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
