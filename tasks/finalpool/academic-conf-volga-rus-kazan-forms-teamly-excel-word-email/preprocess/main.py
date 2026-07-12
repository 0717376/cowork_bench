"""
Preprocess for academic-conf-volga-rus-kazan-forms-teamly-excel-word-email.

Готовит окружение:
- Очищает gform.*, email.*, и пользовательские страницы teamly.
- Создаёт форму «Заявки на конференцию ДРВБ-2026» с 6 вопросами.
- Засевает 10 ответов (7 accept, 3 reject по будущим оценкам рецензентов).
- Создаёт пространство teamly REVIEWS и в нём страницу «Рецензии ДРВБ-2026»
  с таблицей оценок (3 рецензента × 10 заявок).
"""
import argparse
import json
import os
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# Полный список заявок: (ФИО, Email, Аффилиация, Название, Тематика, Аннотация, [3 балла])
SUBMISSIONS = [
    ("Хабибуллин М.Х.", "habibullin@kfu.ru", "Казанский федеральный университет",
     "Археология средневекового Болгара: новые раскопки 2024–2025 гг.",
     "Археология",
     "В докладе представлены результаты двухгодичных полевых работ на территории Болгарского городища; описаны находки IX–X вв. (керамика, нумизматика, фрагменты деревянных конструкций).",
     [5, 5, 4]),
    ("Дроздова О.В.", "drozdova@iling.spb.ru", "ИЛИ РАН, Санкт-Петербург",
     "Берестяные грамоты XII века: новые находки в Новгороде",
     "Источниковедение",
     "Обзор берестяных грамот, найденных при раскопках Троицкого раскопа в 2023–2024 гг.; уточнение датировок и палеографический анализ.",
     [5, 4, 4]),
    ("Петров А.С.", "petrov@hist.msu.ru", "Исторический факультет МГУ",
     "Княжеские договоры XIII века как источник по политической истории",
     "История",
     "Анализируются договорные грамоты Александра Невского, Ярослава Ярославича и Дмитрия Александровича; реконструируется механика княжеских соглашений.",
     [4, 4, 4]),
    ("Шарифуллина Е.Р.", "sharifullina@kfu.ru", "Казанский федеральный университет",
     "Нумизматика волжских булгар: дирхемы и подражания",
     "Археология",
     "Каталог волжско-булгарских монетных находок 2010–2025 гг. Изотопный анализ серебра и реконструкция торговых связей с Багдадским халифатом.",
     [5, 5, 5]),
    ("Иванов К.М.", "ivanov@nov-arch.ru", "Новгородский гос. университет",
     "Софийский собор: археология подземных слоёв",
     "Археология",
     "Результаты георадарного исследования и пробных раскопок под Софийским собором Великого Новгорода. Выявлены три горизонта застройки XI–XIV вв.",
     [5, 4, 4]),
    ("Ласкина Н.Ю.", "laskina@inion.ru", "ИНИОН РАН, Москва",
     "Древнерусская летопись и булгарская хроника: сравнительный анализ",
     "Источниковедение",
     "Сопоставление Лаврентьевской летописи и сообщений Ибн-Фадлана о булгарах. Текстологические параллели и противоречия.",
     [4, 5, 4]),
    ("Соколов Д.А.", "sokolov@kfu.ru", "Казанский федеральный университет",
     "Торговые пути Волги в IX–XI вв.: археологические свидетельства",
     "История",
     "Систематизация находок арабских дирхемов и византийских изделий на Волжском пути. Картирование основных торговых узлов.",
     [5, 4, 5]),
    # 3 отклонённых
    ("Краснов В.П.", "krasnov@example.ru", "Тверской частный исследовательский институт",
     "Тайна копья Святого Олега",
     "История",
     "Гипотеза о происхождении и сакральном значении копья князя Олега. Реконструкция по фольклорным источникам.",
     [2, 3, 3]),
    ("Михайлов С.К.", "mikhailov@yandex.ru", "Независимый исследователь",
     "Криптокурс по Древней Руси",
     "Иное",
     "Авторская трактовка криптографических символов в граффити Софийского собора. Реконструкция «утерянной» системы письма.",
     [1, 2, 2]),
    ("Орлов Р.Е.", "orlov@rambler.ru", "Уфимский центр альтернативной истории",
     "Космогония булгар: версии и факты",
     "Иное",
     "Анализ космогонических представлений булгар через призму современных эзотерических концепций.",
     [3, 2, 3]),
]


def clear_tables(conn):
    """Очистка gform.*, email.*, teamly.* (только пользовательские страницы)."""
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        # Teamly: пользовательские страницы (id > 3 — сидовых страниц 3 в zzz_teamly_after_init.sql)
        try:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        except Exception:
            pass
        try:
            cur.execute("DELETE FROM teamly.spaces WHERE id > 2")
        except Exception:
            pass
        # Email
        for tbl in ("email.attachments", "email.sent_log", "email.drafts"):
            try:
                cur.execute(f"DELETE FROM {tbl}")
            except Exception:
                pass
        cur.execute("DELETE FROM email.messages")
    conn.commit()
    print("[preprocess] Очистка gform/email/teamly выполнена.")


def seed_form(conn):
    """Создаёт форму, 6 вопросов и 10 ответов."""
    form_id = "form_drvb2026"
    with conn.cursor() as cur:
        # Форма
        cur.execute("""
            INSERT INTO gform.forms (id, title, document_title, description, responder_uri, revision_id)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            form_id,
            "Заявки на конференцию ДРВБ-2026",
            "Заявки на конференцию ДРВБ-2026",
            "Сбор заявок на конференцию «Древняя Русь и Волжская Булгария: история и археология» (Казань, 17-19 марта 2026). Приём заявок до 28.02.2026.",
            "https://forms.example.ru/drvb2026",
            "00000001",
        ))

        # Вопросы
        questions = [
            ("ФИО автора", "textQuestion", True, "{}"),
            ("Email", "textQuestion", True, "{}"),
            ("Аффилиация", "textQuestion", True, "{}"),
            ("Название доклада", "textQuestion", True, "{}"),
            ("Тематика", "choiceQuestion", True, json.dumps({"type": "RADIO", "options": [
                {"value": "Археология"}, {"value": "Источниковедение"},
                {"value": "История"}, {"value": "Иное"}
            ]})),
            ("Аннотация", "textQuestion", True, "{}"),
        ]
        for idx, (title, qtype, req, cfg) in enumerate(questions):
            cur.execute("""
                INSERT INTO gform.questions (form_id, title, question_type, required, config, position)
                VALUES (%s, %s, %s, %s, %s::jsonb, %s)
            """, (form_id, title, qtype, req, cfg, idx))

        # Ответы (flat-map в JSONB)
        for idx, (fio, email_, aff, dokl, them, ann, _scores) in enumerate(SUBMISSIONS):
            answers = {
                "ФИО автора": fio,
                "Email": email_,
                "Аффилиация": aff,
                "Название доклада": dokl,
                "Тематика": them,
                "Аннотация": ann,
            }
            cur.execute("""
                INSERT INTO gform.responses (form_id, respondent_email, answers)
                VALUES (%s, %s, %s::jsonb)
            """, (form_id, email_, json.dumps(answers, ensure_ascii=False)))
    conn.commit()
    print(f"[preprocess] Засеяна форма {form_id}: 6 вопросов + {len(SUBMISSIONS)} ответов.")


def seed_teamly_reviews(conn):
    """Создаёт пространство REVIEWS и страницу с таблицей рецензий."""
    rows = []
    for i, (fio, _e, _a, dokl, _t, _ann, scores) in enumerate(SUBMISSIONS, start=1):
        rows.append(f"| {i} | {fio} | {dokl} | {scores[0]} | {scores[1]} | {scores[2]} |")

    body = (
        "# Рецензии ДРВБ-2026\n\n"
        "Оценки рецензентов (шкала 0-5). Среднюю считает оргкомитет.\n\n"
        "| ID | Автор | Название доклада | Рецензент 1 | Рецензент 2 | Рецензент 3 |\n"
        "|---|---|---|---|---|---|\n"
        + "\n".join(rows)
        + "\n\nПорог акцепта: средняя >= 4.0."
    )

    with conn.cursor() as cur:
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('REVIEWS', 'Рецензии конференций', 'Закрытое пространство программного комитета')
            RETURNING id
        """)
        space_id = cur.fetchone()[0]
        cur.execute("""
            INSERT INTO teamly.pages (space_id, title, body, author)
            VALUES (%s, 'Рецензии ДРВБ-2026', %s, 'reviewer-board')
        """, (space_id, body))
    conn.commit()
    print(f"[preprocess] Создано teamly-пространство REVIEWS (id={space_id}) со страницей рецензий.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        seed_form(conn)
        seed_teamly_reviews(conn)
    finally:
        conn.close()

    print("[preprocess] Подготовка academic-conf-volga-rus-kazan завершена.")


if __name__ == "__main__":
    main()
