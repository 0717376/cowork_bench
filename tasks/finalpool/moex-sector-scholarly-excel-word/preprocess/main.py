"""Preprocess script for moex-sector-scholarly-excel-word.

Финансовые данные moex.* засеяны глобально (db/zzz_moex_after_init.sql) и
доступны только для чтения — инъекция/очистка для них не требуется.

Здесь мы готовим только схему scholarly.*: чистим её идемпотентно и
инъектируем релевантные финансовые статьи (sector rotation / industry
analysis / market cycles) плюс шум. Файлы-ответы (xlsx/docx/py/json) НЕ
создаём — их должен сформировать агент.
"""
import os
import argparse, json, os, sys, shutil, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM scholarly.scholar_papers")
    cur.execute("DELETE FROM scholarly.arxiv_papers")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    # Релевантные финансовые статьи (английские заголовки сохранены — eval
    # сопоставляет их как идентификаторы).
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2310.10001', 'Sector Rotation Patterns in Equity Markets', '[{"name": "Author A"}]'::jsonb,
        'An empirical study of cyclical sector rotation across economic regimes, finding rotation cycles every 3-5 years.',
        '["q-fin.PM", "econ.GN"]'::jsonb, 'q-fin.PM', 'https://arxiv.org/pdf/2310.10001', '2023-10-15')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2311.10002', 'Industry Momentum and Cross-Sectional Returns', '[{"name": "Author B"}]'::jsonb,
        'Analysis of industry-level momentum showing persistence of 6-12 months in sector returns.',
        '["q-fin.PM"]'::jsonb, 'q-fin.PM', 'https://arxiv.org/pdf/2311.10002', '2023-11-20')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2312.10003', 'Market Cycles and Sector Performance Dispersion', '[{"name": "Author C"}]'::jsonb,
        'Examines how market cycles drive dispersion in sector performance and volatility.',
        '["q-fin.ST"]'::jsonb, 'q-fin.ST', 'https://arxiv.org/pdf/2312.10003', '2023-12-10')""")
    # Шумовые статьи (нерелевантные)
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2301.00001', 'LLM Reasoning Survey', '[{"name": "Author D"}]'::jsonb, 'A survey of reasoning methods in large language models.',
        '["cs.CL", "cs.AI"]'::jsonb, 'cs.CL', 'https://arxiv.org/pdf/2301.00001', '2023-01-15')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2304.99901', 'Quantum Computing Basics', '[{"name": "Author E"}]'::jsonb, 'Introduction to quantum computing.',
        '["quant-ph"]'::jsonb, 'quant-ph', 'https://arxiv.org/pdf/2304.99901', '2023-04-01')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2305.99902', 'Ocean Modeling Techniques', '[{"name": "Author F"}]'::jsonb, 'Advanced ocean modeling.',
        '["physics.ao-ph"]'::jsonb, 'physics.ao-ph', 'https://arxiv.org/pdf/2305.99902', '2023-05-15')""")
    conn.commit()
    cur.close()
    conn.close()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)

if __name__ == "__main__":
    main()
