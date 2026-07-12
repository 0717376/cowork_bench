"""
Preprocess for arxiv-team-reading-gform-gcal task (RU stack: forms + google_calendar).

Injects ONLY the 7 source papers into scholarly.arxiv_papers,
scholarly.scholar_papers and arxiv.papers:
  - 5 LLM reasoning papers
  - 2 noise papers (different topics)

It does NOT pre-seed the agent's deliverables: the target form
('Reading Group Paper Selection') and the four 'Reading Group Session' calendar
events MUST be created by the agent via the forms / google_calendar MCPs.
Preprocess only clears gform.* and gcal.* target tables for idempotency.

gform.* schema is the RU forms-mcp backend (seeded globally in
db/zzz_gform_after_init.sql). google_calendar is kept as foreign infra; only the
data layer is localized.

Prerequisites:
  - PostgreSQL cowork_gym database running on localhost:5432
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

PAPERS = [
    {
        "arxiv_id": "2401.12345",
        "title": "Chain-of-Thought Reasoning in Large Language Models: A Survey",
        "authors": [{"name": "Wei Zhang"}, {"name": "Li Chen"}],
        "abstract": "We survey recent advances in chain-of-thought reasoning in large language models. Chain-of-thought prompting enables step-by-step reasoning that dramatically improves performance on complex tasks. We analyze over 100 papers and identify key trends in prompt design, model scaling, and evaluation methodologies.",
        "categories": ["cs.CL", "cs.AI"],
        "primary_category": "cs.CL",
        "published": "2024-01-15",
        "citation_count": 450,
        "topic": "LLM_Reasoning",
    },
    {
        "arxiv_id": "2402.23456",
        "title": "Self-Consistency Improves Chain of Thought Reasoning",
        "authors": [{"name": "Xuezhi Wang"}, {"name": "Jason Wei"}],
        "abstract": "We introduce self-consistency as a decoding strategy for chain-of-thought prompting. Instead of greedy decoding, we sample diverse reasoning paths and select the most consistent answer. This approach substantially improves performance on arithmetic and commonsense reasoning benchmarks.",
        "categories": ["cs.CL", "cs.AI"],
        "primary_category": "cs.CL",
        "published": "2024-02-10",
        "citation_count": 380,
        "topic": "LLM_Reasoning",
    },
    {
        "arxiv_id": "2403.34567",
        "title": "Tree of Thoughts: Deliberate Problem Solving with LLMs",
        "authors": [{"name": "Shunyu Yao"}, {"name": "Dian Yu"}],
        "abstract": "We introduce the tree of thoughts framework for deliberate problem solving with large language models. This approach generalizes chain-of-thought by exploring multiple reasoning paths in a tree structure, enabling backtracking and systematic exploration of solution spaces.",
        "categories": ["cs.CL", "cs.AI"],
        "primary_category": "cs.CL",
        "published": "2024-03-20",
        "citation_count": 290,
        "topic": "LLM_Reasoning",
    },
    {
        "arxiv_id": "2404.45678",
        "title": "Least-to-Most Prompting for Complex Reasoning Tasks",
        "authors": [{"name": "Denny Zhou"}, {"name": "Nathanael Scharli"}],
        "abstract": "We propose least-to-most prompting as a technique to enable large language models to solve complex problems by decomposing them into simpler subproblems. The approach achieves state-of-the-art performance on symbolic reasoning, compositional generalization, and math word problems.",
        "categories": ["cs.CL", "cs.AI"],
        "primary_category": "cs.CL",
        "published": "2024-04-05",
        "citation_count": 220,
        "topic": "LLM_Reasoning",
    },
    {
        "arxiv_id": "2405.56789",
        "title": "ReAct: Synergizing Reasoning and Acting in LLMs",
        "authors": [{"name": "Shunyu Yao"}, {"name": "Jeffrey Zhao"}],
        "abstract": "We present ReAct, a paradigm that combines reasoning and acting in large language models. By interleaving reasoning traces with action steps, ReAct enables models to dynamically adjust plans, interact with external environments, and solve complex tasks requiring multi-step decision making.",
        "categories": ["cs.CL", "cs.AI"],
        "primary_category": "cs.CL",
        "published": "2024-05-12",
        "citation_count": 510,
        "topic": "LLM_Reasoning",
    },
    {
        "arxiv_id": "2406.67890",
        "title": "Efficient Training of Vision Transformers",
        "authors": [{"name": "Kai Chen"}],
        "abstract": "We propose efficient training methods for Vision Transformers (ViT) that reduce computation while maintaining accuracy. Our approach combines gradient checkpointing, mixed-precision training, and novel attention approximations to achieve 3x speedup on standard image classification benchmarks.",
        "categories": ["cs.CV", "cs.LG"],
        "primary_category": "cs.CV",
        "published": "2024-06-08",
        "citation_count": 150,
        "topic": "Other",
    },
    {
        "arxiv_id": "2407.78901",
        "title": "Medical Image Segmentation with Deep Learning",
        "authors": [{"name": "Sara Kim"}],
        "abstract": "We apply deep learning methods to medical image segmentation tasks. Our architecture combines U-Net with transformer attention mechanisms to achieve state-of-the-art performance on CT and MRI segmentation benchmarks. The approach shows particular strength on rare pathology detection.",
        "categories": ["cs.CV", "eess.IV"],
        "primary_category": "cs.CV",
        "published": "2024-07-22",
        "citation_count": 89,
        "topic": "Other",
    },
]

def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM scholarly.arxiv_papers")
        cur.execute("DELETE FROM scholarly.scholar_papers")
        cur.execute("DELETE FROM arxiv.papers")
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Cleared all target tables.")


def inject_scholarly_papers(conn):
    with conn.cursor() as cur:
        for p in PAPERS:
            cur.execute("""
                INSERT INTO scholarly.arxiv_papers
                (id, title, authors, abstract, categories, primary_category,
                 published, updated, pdf_url, html_url)
                VALUES (%s, %s, %s::jsonb, %s, %s::jsonb, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    title = EXCLUDED.title,
                    authors = EXCLUDED.authors,
                    abstract = EXCLUDED.abstract
            """, (
                p["arxiv_id"], p["title"], json.dumps(p["authors"]),
                p["abstract"], json.dumps(p["categories"]), p["primary_category"],
                p["published"], p["published"],
                f"http://arxiv.org/pdf/{p['arxiv_id']}",
                f"http://arxiv.org/abs/{p['arxiv_id']}",
            ))
    conn.commit()
    print(f"[preprocess] Injected {len(PAPERS)} papers into scholarly.arxiv_papers")


def inject_scholar_papers(conn):
    # Same 7 papers in scholarly.scholar_papers so all paper surfaces agree
    with conn.cursor() as cur:
        for p in PAPERS:
            pub_year = int(p["published"][:4])
            cur.execute("""
                INSERT INTO scholarly.scholar_papers
                (title, authors, abstract, pub_year, venue, citation_count,
                 url, eprint_url, pub_url, bib)
                VALUES (%s, %s::jsonb, %s, %s, %s, %s, %s, %s, %s, %s::jsonb)
            """, (
                p["title"], json.dumps(p["authors"]), p["abstract"],
                pub_year, "arXiv", p["citation_count"],
                f"http://arxiv.org/abs/{p['arxiv_id']}",
                f"http://arxiv.org/pdf/{p['arxiv_id']}",
                f"http://arxiv.org/abs/{p['arxiv_id']}",
                json.dumps({"year": pub_year, "title": p["title"]}),
            ))
    conn.commit()
    print(f"[preprocess] Injected {len(PAPERS)} papers into scholarly.scholar_papers")


def inject_arxiv_papers(conn):
    with conn.cursor() as cur:
        for p in PAPERS:
            cur.execute("""
                INSERT INTO arxiv.papers
                (id, title, authors, summary, categories, primary_category,
                 pdf_url, published, is_downloaded)
                VALUES (%s, %s, %s::jsonb, %s, %s::jsonb, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    title = EXCLUDED.title,
                    authors = EXCLUDED.authors,
                    summary = EXCLUDED.summary
            """, (
                p["arxiv_id"], p["title"], json.dumps(p["authors"]),
                p["abstract"], json.dumps(p["categories"]), p["primary_category"],
                f"http://arxiv.org/pdf/{p['arxiv_id']}",
                p["published"], True,
            ))
    conn.commit()
    print(f"[preprocess] Injected {len(PAPERS)} papers into arxiv.papers")


def ensure_email_folder(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if not row:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            conn.commit()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        inject_scholarly_papers(conn)
        inject_scholar_papers(conn)
        inject_arxiv_papers(conn)
        ensure_email_folder(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
