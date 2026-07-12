"""Preprocess for terminal-moex-scholarly-excel-word-email.
Инъекция научных статей. Очистка почты. Данные moex.* засеяны глобально
(схема moex.*, db/zzz_moex_after_init.sql) и доступны только для чтения."""
import argparse
import glob as globmod
import json
import os
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Clear email data
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.drafts")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared email data.")

        # Clear scholarly data
        cur.execute("DELETE FROM scholarly.scholar_papers")
        conn.commit()
        print("[preprocess] Cleared scholarly papers.")

        # Inject relevant scholarly papers
        # Канонические названия и авторы статей сохранены на английском
        # (eval сопоставляет их по подстроке); русифицированы только аннотации.
        papers = [
            {
                "title": "Efficient Capital Markets: A Review of Theory and Empirical Work",
                "authors": json.dumps(["Eugene Fama"]),
                "abstract": "Обзор теоретической и эмпирической литературы по гипотезе эффективного "
                           "рынка. Работа классифицирует исследования на три категории: тесты слабой, "
                           "полусильной и сильной форм эффективности рынка.",
                "pub_year": 1970,
                "venue": "Journal of Finance",
                "citation_count": 15000,
            },
            {
                "title": "Random Walks in Stock Market Prices",
                "authors": json.dumps(["Eugene Fama"]),
                "abstract": "Исследование модели случайного блуждания и её следствий для поведения "
                           "цен на фондовом рынке. Статья приводит эмпирические свидетельства того, "
                           "что цены акций приблизительно следуют случайному блужданию.",
                "pub_year": 1965,
                "venue": "Financial Analysts Journal",
                "citation_count": 8500,
            },
            {
                "title": "Algorithmic Trading and Market Quality",
                "authors": json.dumps(["Terrence Hendershott", "Charles Jones", "Albert Menkveld"]),
                "abstract": "Исследование влияния алгоритмической торговли на качество рынка, включая "
                           "ликвидность, формирование цены и волатильность. Результаты показывают, что "
                           "алгоритмическая торговля повышает эффективность рынка.",
                "pub_year": 2011,
                "venue": "Journal of Finance",
                "citation_count": 3200,
            },
            {
                "title": "Market Efficiency in the Age of Big Data",
                "authors": json.dumps(["David Easley", "Marcos Lopez de Prado", "Maureen O'Hara"]),
                "abstract": "Рассматривается, как современная аналитика данных, машинное обучение и "
                           "альтернативные источники данных влияют на гипотезу эффективного рынка и "
                           "механизмы формирования цены.",
                "pub_year": 2021,
                "venue": "Journal of Financial Economics",
                "citation_count": 850,
            },
            {
                "title": "Weak-Form Efficiency of the Russian Stock Market (MOEX)",
                "authors": json.dumps(["A. Bukhvalov", "S. Smirnov"]),
                "abstract": "Эмпирическая проверка слабой формы эффективности российского фондового "
                           "рынка на примере индекса Московской биржи (MOEX) и голубых фишек, включая "
                           "Сбербанк и Лукойл. Тесты на автокорреляцию и нормальность доходностей.",
                "pub_year": 2018,
                "venue": "Журнал Новой экономической ассоциации",
                "citation_count": 120,
            },
        ]

        # Шумовые статьи (нерелевантные)
        noise_papers = [
            {
                "title": "Deep Learning for Natural Language Processing: A Survey",
                "authors": json.dumps(["Various Authors"]),
                "abstract": "Обзор методов глубокого обучения применительно к задачам обработки "
                           "естественного языка.",
                "pub_year": 2020,
                "venue": "ACL",
                "citation_count": 500,
            },
            {
                "title": "Climate Change and Agricultural Productivity",
                "authors": json.dumps(["Smith et al."]),
                "abstract": "Анализ влияния изменения климата на урожайность сельскохозяйственных культур.",
                "pub_year": 2019,
                "venue": "Nature",
                "citation_count": 300,
            },
            {
                "title": "Quantum Computing Applications in Cryptography",
                "authors": json.dumps(["Chen et al."]),
                "abstract": "Обзор применения квантовых вычислений в современной криптографии.",
                "pub_year": 2022,
                "venue": "IEEE",
                "citation_count": 200,
            },
        ]

        for i, paper in enumerate(papers + noise_papers):
            cur.execute("""
                INSERT INTO scholarly.scholar_papers (id, title, authors, abstract, pub_year, venue, citation_count)
                VALUES (%s, %s, %s::jsonb, %s, %s, %s, %s)
            """, (i + 1, paper["title"], paper["authors"], paper["abstract"],
                  paper["pub_year"], paper["venue"], paper["citation_count"]))

        conn.commit()
        print(f"[preprocess] Injected {len(papers)} relevant + {len(noise_papers)} noise scholarly papers.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Clean up agent workspace
    if args.agent_workspace:
        for pattern in ["FinTech_Research_Report.xlsx", "FinTech_Research_Report.docx", "market_analysis.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
