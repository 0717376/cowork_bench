"""Preprocess script for canvas-scholarly-curriculum-excel-teamly.

Swaps notion -> teamly (RU Confluence-analog, schema teamly.*). Seeds the
scholarly tables with papers whose topics (machine learning / data analytics /
computational thinking) and citation counts are self-consistent, so the
agent's Research_Trends sheet is reproducible by the evaluator from the same
source. Injects a single RU noise page in the target Teamly space so the agent
has leftover content to ignore -- it must NOT be the deliverable.

Does NOT pre-seed the 'Curriculum Innovation Tracker' page, the xlsx, or any
JSON deliverable. Canvas data is provided by the globally seeded canvas.*
schema and is read live by both the agent and the evaluator.
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

# Relevant scholarly papers, on-topic for the curriculum (machine learning /
# data analytics / computational thinking) with citation counts. English
# identifiers (titles/topics) are preserved because the eval matches them as
# substrings. These are the reproducible source for the Research_Trends sheet.
RELEVANT_PAPERS = [
    {
        "id": "2401.10001",
        "title": "Deep Learning Advances in Machine Learning",
        "topic": "machine learning",
        "abstract": "A survey of recent advances in deep learning for machine learning tasks.",
        "categories": ["cs.LG"], "primary": "cs.LG",
        "year": 2024, "venue": "NeurIPS 2024", "citations": 150,
    },
    {
        "id": "2402.10002",
        "title": "Scalable Data Analytics Methods",
        "topic": "data analytics",
        "abstract": "Methods for scalable data analytics over large datasets.",
        "categories": ["cs.DB", "cs.LG"], "primary": "cs.DB",
        "year": 2024, "venue": "VLDB 2024", "citations": 120,
    },
    {
        "id": "2403.10003",
        "title": "Computational Thinking in Education",
        "topic": "computational thinking",
        "abstract": "Integrating computational thinking into the computer science curriculum.",
        "categories": ["cs.CY"], "primary": "cs.CY",
        "year": 2023, "venue": "SIGCSE 2023", "citations": 80,
    },
]

# Noise papers (off-topic) the agent must ignore.
NOISE_PAPERS = [
    {
        "id": "2304.99901",
        "title": "Quantum Computing Basics",
        "topic": "quantum computing",
        "abstract": "Introduction to quantum computing.",
        "categories": ["quant-ph"], "primary": "quant-ph",
        "year": 2023, "venue": "arXiv preprint", "citations": 12,
    },
    {
        "id": "2305.99902",
        "title": "Ocean Modeling Techniques",
        "topic": "ocean modeling",
        "abstract": "Advanced ocean modeling.",
        "categories": ["physics.ao-ph"], "primary": "physics.ao-ph",
        "year": 2023, "venue": "arXiv preprint", "citations": 9,
    },
]

ALL_PAPERS = RELEVANT_PAPERS + NOISE_PAPERS

TEAMLY_SPACE_KEY = "CURRICULUM"


def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)


def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    # Teamly: drop user-created pages (seed pages have id <= 3); ensure a space.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES (%s, 'Учебная программа',
                        'База знаний кафедры по обзору и развитию учебной программы.')
                ON CONFLICT (key) DO NOTHING
            """, (TEAMLY_SPACE_KEY,))
    except Exception as e:
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")
    cur.execute("DELETE FROM scholarly.scholar_papers")
    cur.execute("DELETE FROM scholarly.arxiv_papers")
    conn.commit()
    cur.close()
    conn.close()


def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()

    # Scholarly papers in both arxiv_papers and scholar_papers tables.
    for p in ALL_PAPERS:
        cur.execute(
            """INSERT INTO scholarly.arxiv_papers
               (id, title, authors, abstract, categories, primary_category, pdf_url, published)
               VALUES (%s, %s, %s::jsonb, %s, %s::jsonb, %s, %s, %s)""",
            (p["id"], p["title"], json.dumps([{"name": "Author " + p["id"][-1]}]),
             p["abstract"], json.dumps(p["categories"]), p["primary"],
             f"https://arxiv.org/pdf/{p['id']}", f"{p['year']}-06-01"),
        )
        cur.execute(
            """INSERT INTO scholarly.scholar_papers
               (title, authors, abstract, pub_year, venue, citation_count, url, eprint_url, pub_url, bib)
               VALUES (%s, %s::jsonb, %s, %s, %s, %s, %s, %s, %s, %s::jsonb)""",
            (p["title"], json.dumps([{"name": "Author " + p["id"][-1]}]),
             p["abstract"], p["year"], p["venue"], p["citations"],
             f"https://arxiv.org/abs/{p['id']}",
             f"https://arxiv.org/pdf/{p['id']}",
             f"https://arxiv.org/abs/{p['id']}",
             json.dumps({"title": p["title"], "year": p["year"]})),
        )

    # Single RU noise page in the target Teamly space. It must NOT satisfy the
    # 'Curriculum Innovation Tracker' deliverable check.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("SELECT id FROM teamly.spaces WHERE key = %s", (TEAMLY_SPACE_KEY,))
            row = cur.fetchone()
            if row is None:
                cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
                row = cur.fetchone()
            space_id = row[0] if row else None
            if space_id is not None:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) "
                    "VALUES (%s, %s, %s, %s)",
                    (space_id, "Архив протоколов совещаний кафедры",
                     "Старые протоколы заседаний кафедры. Не относится к текущему обзору учебной программы.",
                     "team"),
                )
    except Exception as e:
        print(f"[preprocess] WARNING: noise teamly page skipped: {e}")

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
