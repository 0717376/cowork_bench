"""Preprocess for terminal-arxiv-scholarly-teamly-word-excel.
Clears arxiv, scholarly schemas and leftover Teamly "Research Paper Tracker"
space/pages (idempotency). Injects 6 transformer papers + 3 noise papers.

We do NOT pre-create the answer space/pages — the agent must build the
"Research Paper Tracker" space and its per-paper pages itself, so the
evaluation actually exercises the agent's work.

The Teamly schema is seeded globally (db/zzz_teamly_after_init.sql); no new
db files are added here.
"""
import argparse
import json

import os
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")

# 6 transformer/NLP papers (appear in both scholarly and arxiv)
TRANSFORMER_PAPERS = [
    {
        "id": "1706.03762",
        "title": "Attention Is All You Need",
        "authors": [{"name": "Ashish Vaswani"}, {"name": "Noam Shazeer"}, {"name": "Niki Parmar"}],
        "published": "2017-06-12T00:00:00",
        "journal_ref": "NeurIPS 2017",
        "summary": "The dominant sequence transduction models are based on complex recurrent or convolutional neural networks. We propose a new simple network architecture, the Transformer, based solely on attention mechanisms, dispensing with recurrence and convolutions entirely.",
        "categories": ["cs.CL", "cs.LG"],
        "category": "Architecture Design"
    },
    {
        "id": "1810.04805",
        "title": "BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding",
        "authors": [{"name": "Jacob Devlin"}, {"name": "Ming-Wei Chang"}, {"name": "Kenton Lee"}],
        "published": "2018-10-11T00:00:00",
        "journal_ref": "NAACL 2019",
        "summary": "We introduce a new language representation model called BERT, designed to pre-train deep bidirectional representations from unlabeled text by jointly conditioning on both left and right context. The pre-trained model can be fine-tuned for various tasks.",
        "categories": ["cs.CL"],
        "category": "Training Methods"
    },
    {
        "id": "2005.14165",
        "title": "Language Models are Few-Shot Learners",
        "authors": [{"name": "Tom B. Brown"}, {"name": "Benjamin Mann"}, {"name": "Nick Ryder"}],
        "published": "2020-05-28T00:00:00",
        "journal_ref": "NeurIPS 2020",
        "summary": "We demonstrate that scaling up language models greatly improves task-agnostic, few-shot performance, sometimes even reaching competitiveness with prior state-of-the-art fine-tuning approaches. We train GPT-3, a Transformer-based language model with 175 billion parameters, and evaluate its performance in the few-shot setting.",
        "categories": ["cs.CL", "cs.LG"],
        "category": "Training Methods"
    },
    {
        "id": "1409.0473",
        "title": "Neural Machine Translation by Jointly Learning to Align and Translate",
        "authors": [{"name": "Dzmitry Bahdanau"}, {"name": "Kyunghyun Cho"}, {"name": "Yoshua Bengio"}],
        "published": "2014-09-01T00:00:00",
        "journal_ref": "ICLR 2015",
        "summary": "Neural machine translation is a recently proposed approach to the machine translation task. We conjecture that the use of a fixed-length vector is a bottleneck and propose to automatically search for parts of a source sentence that are relevant to predicting a target word. This soft attention markedly improves translation quality on standard benchmark translation tasks, with the largest gains on long sentences.",
        "categories": ["cs.CL", "cs.LG"],
        "category": "Applications"
    },
    {
        "id": "1910.10683",
        "title": "Exploring the Limits of Transfer Learning with a Unified Text-to-Text Transformer",
        "authors": [{"name": "Colin Raffel"}, {"name": "Noam Shazeer"}, {"name": "Adam Roberts"}],
        "published": "2019-10-23T00:00:00",
        "journal_ref": "JMLR 2020",
        "summary": "We introduce T5, an encoder-decoder Transformer architecture whose unified design converts all text-based language problems into a text-to-text format. The architecture of the model -- its layer structure and attention mechanism -- is the central contribution of this work.",
        "categories": ["cs.CL", "cs.LG"],
        "category": "Architecture Design"
    },
    {
        "id": "2009.06732",
        "title": "Efficient Transformers: A Survey",
        "authors": [{"name": "Yi Tay"}, {"name": "Mostafa Dehghani"}, {"name": "Dara Bahri"}],
        "published": "2020-09-14T00:00:00",
        "journal_ref": "ACM Computing Surveys 2022",
        "summary": "Transformers have been the driving force behind many recent advances in NLP. This paper provides a comprehensive survey and review of efficient transformer architectures, giving an overview of methods that reduce the computational and memory complexity of transformers.",
        "categories": ["cs.CL", "cs.LG"],
        "category": "Survey"
    },
]

# 3 noise papers (not about transformers, only in arxiv)
NOISE_PAPERS = [
    {
        "id": "1207.00580",
        "title": "Improving Neural Networks by Preventing Co-adaptation of Feature Detectors",
        "authors": [{"name": "Geoffrey E. Hinton"}, {"name": "Nitish Srivastava"}],
        "published": "2012-07-03T00:00:00",
        "summary": "We describe a technique called dropout for addressing overfitting in neural networks.",
        "categories": ["cs.NE", "cs.LG"]
    },
    {
        "id": "1502.03167",
        "title": "Batch Normalization: Accelerating Deep Network Training",
        "authors": [{"name": "Sergey Ioffe"}, {"name": "Christian Szegedy"}],
        "published": "2015-02-11T00:00:00",
        "summary": "Training deep neural networks is complicated by the fact that the distribution of each layer inputs changes during training. We propose batch normalization.",
        "categories": ["cs.LG"]
    },
    {
        "id": "1312.06199",
        "title": "Playing Atari with Deep Reinforcement Learning",
        "authors": [{"name": "Volodymyr Mnih"}, {"name": "Koray Kavukcuoglu"}],
        "published": "2013-12-19T00:00:00",
        "summary": "We present the first deep learning model to successfully learn control policies directly from high-dimensional sensory input using reinforcement learning.",
        "categories": ["cs.LG", "cs.AI"]
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
        # Clear source schemas (scholarly + arxiv).
        cur.execute("DELETE FROM scholarly.arxiv_papers")
        cur.execute("DELETE FROM scholarly.scholar_papers")
        cur.execute("DELETE FROM arxiv.papers")
        conn.commit()
        print("[preprocess] Cleared scholarly, arxiv schemas.")

        # Idempotency for Teamly: remove any leftover "Research Paper Tracker"
        # space (cascades to its pages) from previous runs. Do NOT pre-create
        # the tracker space or paper pages — that is the agent's deliverable.
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is None:
            print("[preprocess] WARNING: teamly schema not found; "
                  "run db/zzz_teamly_after_init.sql.")
        else:
            cur.execute(
                "DELETE FROM teamly.spaces "
                "WHERE name ILIKE '%research paper%' OR name ILIKE '%paper tracker%' "
                "OR key = 'RPT'"
            )
            # Also clear stray paper pages that may have landed in seed spaces.
            cur.execute(
                "DELETE FROM teamly.pages WHERE title ILIKE '%paper tracker%'"
            )
            conn.commit()
            print("[preprocess] Cleared leftover Teamly 'Research Paper Tracker' space/pages.")

        # Inject scholarly papers (6 transformer papers)
        for p in TRANSFORMER_PAPERS:
            cur.execute("""
                INSERT INTO scholarly.arxiv_papers (id, title, authors, published, journal_ref, abstract)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    title = EXCLUDED.title, authors = EXCLUDED.authors,
                    published = EXCLUDED.published, journal_ref = EXCLUDED.journal_ref,
                    abstract = EXCLUDED.abstract
            """, (p["id"], p["title"], json.dumps(p["authors"]),
                  p["published"], p["journal_ref"], p["summary"]))
        conn.commit()
        print(f"[preprocess] Injected {len(TRANSFORMER_PAPERS)} papers into scholarly.arxiv_papers")

        # Inject the same 6 transformer papers into scholarly.scholar_papers so
        # the obvious "Google Scholar" tools (search-google-scholar /
        # list-scholar-papers) surface them. The arxiv id is embedded in the
        # url fields so the agent can establish the cross-source overlap.
        for p in TRANSFORMER_PAPERS:
            cur.execute("""
                INSERT INTO scholarly.scholar_papers
                (title, authors, abstract, pub_year, venue, citation_count,
                 url, eprint_url, pub_url, bib)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                p["title"], json.dumps(p["authors"]), p["summary"],
                int(p["published"][:4]), p["journal_ref"], 0,
                f"http://arxiv.org/abs/{p['id']}",
                f"http://arxiv.org/pdf/{p['id']}",
                f"http://arxiv.org/abs/{p['id']}",
                json.dumps({"title": p["title"], "year": int(p["published"][:4])}),
            ))
        conn.commit()
        print(f"[preprocess] Injected {len(TRANSFORMER_PAPERS)} papers into scholarly.scholar_papers")

        # Inject arxiv papers (6 transformer + 3 noise)
        all_arxiv = TRANSFORMER_PAPERS + NOISE_PAPERS
        for p in all_arxiv:
            cur.execute("""
                INSERT INTO arxiv.papers (id, title, authors, summary, published, categories)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    title = EXCLUDED.title, authors = EXCLUDED.authors,
                    summary = EXCLUDED.summary, published = EXCLUDED.published,
                    categories = EXCLUDED.categories
            """, (p["id"], p["title"], json.dumps(p["authors"]),
                  p["summary"], p["published"], json.dumps(p["categories"])))
        conn.commit()
        print(f"[preprocess] Injected {len(all_arxiv)} papers into arxiv.papers")

        # Verify
        cur.execute("SELECT COUNT(*) FROM scholarly.arxiv_papers")
        s_count = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM arxiv.papers")
        a_count = cur.fetchone()[0]
        print(f"[preprocess] scholarly: {s_count} papers, arxiv: {a_count} papers")

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
