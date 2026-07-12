"""Preprocess for terminal-arxiv-moex-excel-word-teamly-email.
Clears arxiv, teamly (user pages), email schemas. Injects 6 papers (4 relevant +
2 noise) into arxiv. Injects noise teamly page + noise email data.
Titles / arxiv IDs / author names stay ENGLISH (eval greps them); summaries are RU.
"""
import argparse
import json
import os
import uuid

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")

# 4 relevant papers (AI + finance). Titles/IDs/authors EN; summaries RU prose.
RELEVANT_PAPERS = [
    {
        "id": "2306.06031",
        "title": "FinGPT: Large Language Models for Financial Applications",
        "authors": [{"name": "Hongyang Yang"}, {"name": "Xiao-Yang Liu"}, {"name": "Christina Dan Wang"}],
        "summary": "Большие языковые модели (LLM) продемонстрировали выдающиеся возможности в различных задачах NLP. В этой работе представлен FinGPT — фреймворк с открытым исходным кодом для финансовых больших языковых моделей. FinGPT предоставляет доступные инструменты для анализа финансовой тональности, прогнозирования движения котировок и автоматической генерации финансовых отчётов с использованием передовых языковых моделей.",
        "categories": ["cs.CL", "q-fin.ST"],
        "primary_category": "cs.CL",
        "published": "2023-06-09",
        "pdf_url": "https://arxiv.org/pdf/2306.06031"
    },
    {
        "id": "2304.07619",
        "title": "Can Large Language Models Predict Stock Price Movements?",
        "authors": [{"name": "Qianqian Chen"}, {"name": "Yuwei Li"}, {"name": "Feng Zhang"}],
        "summary": "В этой работе исследуется, способны ли большие языковые модели, такие как ChatGPT, прогнозировать движения котировок акций. Мы строим эксперименты на основе заголовков финансовых новостей и оцениваем качество прогнозов LLM в режимах zero-shot и few-shot. Результаты показывают, что LLM достигают многообещающей точности в предсказании направления движения акций, превосходя традиционные подходы анализа тональности.",
        "categories": ["q-fin.ST", "cs.CL"],
        "primary_category": "q-fin.ST",
        "published": "2023-04-15",
        "pdf_url": "https://arxiv.org/pdf/2304.07619"
    },
    {
        "id": "2302.14040",
        "title": "Deep Learning for Financial Risk Prediction",
        "authors": [{"name": "Rajesh Kumar"}, {"name": "Priya Patel"}, {"name": "Amit Singh"}],
        "summary": "Мы предлагаем фреймворк глубокого обучения для оценки кредитного риска и прогнозирования финансовых рисков. Наша модель сочетает рекуррентные нейронные сети с механизмами внимания для обработки последовательных финансовых данных. Эксперименты на банковских наборах данных показывают значительные улучшения по сравнению с традиционными статистическими моделями в прогнозировании дефолтов по кредитам и оценок кредитного риска.",
        "categories": ["q-fin.RM", "cs.LG"],
        "primary_category": "q-fin.RM",
        "published": "2023-02-27",
        "pdf_url": "https://arxiv.org/pdf/2302.14040"
    },
    {
        "id": "2311.10723",
        "title": "Machine Learning in Quantitative Finance: Applications and Challenges",
        "authors": [{"name": "Carlos Martinez"}, {"name": "Wei Zhao"}, {"name": "David Brown"}],
        "summary": "Этот обзорный труд рассматривает применение машинного обучения в количественных финансах, охватывая алгоритмическую торговлю, оптимизацию портфеля, управление рисками и анализ микроструктуры рынка. Мы изучаем, как глубокое обучение, обучение с подкреплением и обработка естественного языка трансформируют принятие финансовых решений, и обсуждаем такие вызовы, как качество данных, интерпретируемость и соответствие регуляторным требованиям.",
        "categories": ["q-fin.CP", "cs.LG"],
        "primary_category": "q-fin.CP",
        "published": "2023-11-17",
        "pdf_url": "https://arxiv.org/pdf/2311.10723"
    },
]

# 2 noise papers (not finance-related). EN titles/IDs/authors; RU summaries.
NOISE_PAPERS = [
    {
        "id": "2305.18290",
        "title": "Drag Your GAN: Interactive Point-based Manipulation on the Generative Image Manifold",
        "authors": [{"name": "Xingang Pan"}, {"name": "Ayush Tewari"}, {"name": "Thomas Leimkuhler"}],
        "summary": "Мы представляем DragGAN — подход, позволяющий пользователям интерактивно перетаскивать содержимое изображения в точные целевые позиции. Эта техника обеспечивает реалистичное редактирование изображений за счёт перемещения точек на генеративном многообразии изображений.",
        "categories": ["cs.CV", "cs.GR"],
        "primary_category": "cs.CV",
        "published": "2023-05-25",
        "pdf_url": "https://arxiv.org/pdf/2305.18290"
    },
    {
        "id": "2307.09288",
        "title": "Llama 2: Open Foundation and Fine-Tuned Chat Models",
        "authors": [{"name": "Hugo Touvron"}, {"name": "Louis Martin"}, {"name": "Kevin Stone"}],
        "summary": "Мы разрабатываем и выпускаем Llama 2 — набор предобученных и дообученных больших языковых моделей с числом параметров от 7B до 70B. Наши дообученные модели, называемые Llama 2-Chat, оптимизированы для диалоговых сценариев и превосходят большинство открытых диалоговых моделей по показателям полезности и безопасности.",
        "categories": ["cs.CL", "cs.AI"],
        "primary_category": "cs.CL",
        "published": "2023-07-18",
        "pdf_url": "https://arxiv.org/pdf/2307.09288"
    },
]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        # Clear schemas
        cur.execute("DELETE FROM arxiv.papers")
        # Teamly: drop user-created pages (seed pages have id <= 3); ensure a space.
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('RESEARCH', 'Исследования',
                        'Корпоративная база знаний количественной аналитики: статьи, конвейеры, инвестиционные тезисы.')
                ON CONFLICT (key) DO NOTHING
            """)
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared arxiv, teamly (user pages), email schemas.")

        # Inject arxiv papers (4 relevant + 2 noise)
        all_papers = RELEVANT_PAPERS + NOISE_PAPERS
        for p in all_papers:
            cur.execute("""
                INSERT INTO arxiv.papers (id, title, authors, summary, categories, primary_category, published, pdf_url, is_downloaded)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, false)
                ON CONFLICT (id) DO UPDATE SET
                    title = EXCLUDED.title, authors = EXCLUDED.authors,
                    summary = EXCLUDED.summary, categories = EXCLUDED.categories,
                    primary_category = EXCLUDED.primary_category,
                    published = EXCLUDED.published, pdf_url = EXCLUDED.pdf_url
            """, (p["id"], p["title"], json.dumps(p["authors"]),
                  p["summary"], json.dumps(p["categories"]),
                  p["primary_category"], p["published"], p["pdf_url"]))
        conn.commit()
        print(f"[preprocess] Injected {len(all_papers)} papers into arxiv.papers")

        # Inject noise teamly page (RU) in the RESEARCH space — leftover content
        # the agent must ignore; must NOT satisfy the Research Pipeline check.
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("SELECT id FROM teamly.spaces WHERE key = 'RESEARCH'")
            row = cur.fetchone()
            if row is None:
                cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
                row = cur.fetchone()
            space_id = row[0] if row else None
            if space_id is not None:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
                    (space_id, "Архив протоколов совещаний",
                     "Старые заметки со встреч команды. Не относится к текущей задаче по конвейеру статей.",
                     "team"),
                )
                conn.commit()
                print("[preprocess] Injected noise teamly page.")

        # Inject noise email data
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            inbox_id = row[0]
        else:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            inbox_id = cur.fetchone()[0]
            cur.execute("INSERT INTO email.folders (name) VALUES ('Sent')")
            cur.execute("INSERT INTO email.folders (name) VALUES ('Drafts')")
            conn.commit()

        noise_emails = [
            ("Напоминание о совещании: планирование Q4", "admin@firm.com", "team@firm.com",
             "Пожалуйста, присоединитесь к совещанию по планированию Q4 завтра в 14:00."),
            ("Заказ обеда на пятницу", "social@firm.com", "office@firm.com",
             "Пожалуйста, отправьте ваши заказы на обед до конца дня в четверг."),
        ]
        for subj, from_a, to_a, body in noise_emails:
            cur.execute("""
                INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                VALUES (%s, %s, %s, %s, %s, NOW(), %s, false)
            """, (inbox_id, str(uuid.uuid4()), subj, from_a, json.dumps([to_a]), body))
        conn.commit()
        print("[preprocess] Injected noise email data.")

        # Verify
        cur.execute("SELECT COUNT(*) FROM arxiv.papers")
        print(f"[preprocess] arxiv.papers: {cur.fetchone()[0]} papers")
        cur.execute("SELECT COUNT(*) FROM email.messages")
        print(f"[preprocess] email.messages: {cur.fetchone()[0]} messages")

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
