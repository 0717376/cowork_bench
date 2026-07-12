#!/usr/bin/env python3
"""
Preprocess for arxiv-research-workflow-pipeline (RU, teamly swap).

- Clears user-created Teamly pages (keeps seed pages id<=3 as format example),
  so the agent builds the research hub from scratch (no pre-seeded answer).
- Injects a consumable RU-described source CSV `paper_sources.csv` into the
  agent workspace. It carries the per-paper bibliographic data + citation counts
  so that paper_analysis.xlsx values are DERIVABLE (no fabricated answer key).
- Clears leftover answer artifacts for idempotency. Does NOT create the answer
  files paper_analysis.xlsx / literature_review.docx (no pre-seeding).
"""
import argparse
import csv
import os
from pathlib import Path

try:
    import psycopg2
except Exception:  # pragma: no cover
    psycopg2 = None

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

# English schema field names are kept verbatim (eval greps them). Prose/desc is RU.
SOURCE_HEADER = ["Paper ID", "Title", "Authors", "Year", "Citation Count", "Category"]

# Consumable source: per-paper bibliographic data with citation counts.
# Values are the legitimate source the agent reads to build paper_analysis.xlsx.
PAPERS = [
    {"id": '2001.03134', "title": 'Federated Machine Learning: Concept and Applications', "authors": 'Konecny, McMahan, Yu', "year": 2020, "citations": 1250, "category": 'Federated Learning'},
    {"id": '2007.14861', "title": 'Communication-Efficient Learning of Deep Networks', "authors": 'McMahan, Moore, Ramage', "year": 2021, "citations": 3200, "category": 'Communication Efficiency'},
    {"id": '2010.12277', "title": 'Towards Federated Learning at Scale: System Design', "authors": 'Bonawitz et al.', "year": 2021, "citations": 2150, "category": 'System Design'},
    {"id": '2009.01974', "title": 'Adaptive Federated Optimization', "authors": 'Reddi, Charles, Zaheer', "year": 2020, "citations": 890, "category": 'Optimization'},
    {"id": '2106.05046', "title": 'Federated Learning with Differential Privacy', "authors": 'Oh, Zhu', "year": 2021, "citations": 520, "category": 'Privacy'},
    {"id": '2102.06701', "title": 'Asynchronous Federated Optimization', "authors": 'Wang, Joshi, Kale', "year": 2021, "citations": 340, "category": 'Asynchronous Methods'},
    {"id": '2005.12394', "title": 'Gradient Compression for Distributed Training', "authors": 'You, Zhang, Demmel', "year": 2020, "citations": 680, "category": 'Compression'},
    {"id": '2109.09813', "title": 'Personalized Federated Learning with Theoretical Guarantees', "authors": 'Fallah, Mokhtari, Ozdaglar', "year": 2021, "citations": 285, "category": 'Personalization'},
    {"id": '2012.06559', "title": 'Field Guide to Federated Optimization', "authors": 'Charles, Konecny', "year": 2020, "citations": 450, "category": 'Survey'},
    {"id": '2006.04638', "title": 'Towards Optimal and Communication-Efficient Distributed ML', "authors": 'Suresh et al.', "year": 2020, "citations": 380, "category": 'Theory'},
    {"id": '2103.01048', "title": 'Federated Unlearning', "authors": 'Liu, Jiang, Parkes', "year": 2021, "citations": 210, "category": 'Unlearning'},
    {"id": '2110.15811', "title": 'Accelerated Federated Learning with Decoupled Optimization', "authors": 'Zhang, Garg, Vassilvitskii', "year": 2021, "citations": 165, "category": 'Acceleration'},
    {"id": '2003.00295', "title": 'A Distributed Approach towards Distance Metric Learning', "authors": 'Zhu, Sun', "year": 2020, "citations": 95, "category": 'Metric Learning'},
    {"id": '2107.02514', "title": 'Byzantine-Resilient Distributed Learning', "authors": 'Gupta, Chandrachoodan, Vaidya', "year": 2021, "citations": 145, "category": 'Byzantine Robustness'},
    {"id": '2105.13029', "title": 'Federated Learning for Edge Computing: A Survey', "authors": 'Wang, Lan, Shroff', "year": 2021, "citations": 280, "category": 'Edge Computing'},
    {"id": '2006.11537', "title": 'Variance Reduction is an Antidote to Byzantine Poisoning', "authors": 'Vogels, Karimireddy, Jaggi', "year": 2020, "citations": 220, "category": 'Robustness'},
    {"id": '2108.12849', "title": 'Communication Efficient Learning with Sparsification', "authors": 'Zhang, Singh, Yang', "year": 2021, "citations": 125, "category": 'Sparsification'},
    {"id": '2104.01155', "title": 'Clustered Federated Learning: Model-Agnostic', "authors": 'Ghosh, Chung, Ding', "year": 2021, "citations": 190, "category": 'Clustering'},
    {"id": '2010.00145', "title": 'Scaffold: Stochastic Controlled Averaging', "authors": 'Karimireddy et al.', "year": 2020, "citations": 310, "category": 'Variance Reduction'},
    {"id": '2109.11214', "title": 'Client Selection in Federated Learning', "authors": 'Wang, Sahu, Joshi', "year": 2021, "citations": 185, "category": 'Client Selection'},
    {"id": '2007.07314', "title": 'Optimal Algorithms for Federated Learning with Heterogeneous Data', "authors": 'Mohri, Sivek', "year": 2020, "citations": 275, "category": 'Heterogeneous Data'},
    {"id": '2102.05098', "title": 'Federated Learning with Non-IID Data', "authors": 'Naamad, Dolev', "year": 2021, "citations": 165, "category": 'Non-IID Data'},
    {"id": '2006.05582', "title": 'Gradient Boosting Machine Learning', "authors": 'Ruan, Vassilvitskii', "year": 2020, "citations": 420, "category": 'Gradient Boosting'},
    {"id": '2108.13572', "title": 'Layer-wise Coordination for Efficient Learning', "authors": 'Marques, Agrawal, Li', "year": 2021, "citations": 110, "category": 'Layer-wise Methods'},
    {"id": '2103.13821', "title": 'Convergence Analysis of Two-Timescale Learning', "authors": 'Hong, Razaviyayn, Luo', "year": 2021, "citations": 135, "category": 'Convergence Theory'},
    {"id": '2005.09258', "title": 'Communication-Efficient Agnostic Federated Averaging', "authors": 'Park, Kairouz, Rabbat', "year": 2020, "citations": 190, "category": 'Averaging Methods'},
    {"id": '2109.05123', "title": 'Federated Learning with Adaptive Cluster Selection', "authors": 'Zhou, Eldar', "year": 2021, "citations": 145, "category": 'Adaptive Clustering'},
    {"id": '2006.17041', "title": 'Throughput-Optimal Topology Design for ML', "authors": 'Ramamoorthy, Kairouz', "year": 2020, "citations": 215, "category": 'Network Design'},
    {"id": '2110.06245', "title": 'Privacy-Preserving Federated Learning with MPC', "authors": 'Yedekar, Dutta, Vaidya', "year": 2021, "citations": 180, "category": 'Cryptography'},
    {"id": '2004.12142', "title": 'Performance Tuning for Distributed ML', "authors": 'Huang, Larus, Lee', "year": 2020, "citations": 310, "category": 'Performance'},
    {"id": '2111.01034', "title": 'Federated Multi-Task Learning with Adaptive Clustering', "authors": 'Rajabi, Rebjock, Januschowski', "year": 2021, "citations": 125, "category": 'Multi-Task'},
    {"id": '2002.11364', "title": 'Communication-Efficient Distributed Deep Learning: Survey', "authors": 'Vogels, Karimireddy, Jaggi', "year": 2020, "citations": 950, "category": 'Survey'},
    {"id": '2105.04821', "title": 'Federated Learning with Heterogeneous Labels', "authors": 'Han, Bi, Wang', "year": 2021, "citations": 140, "category": 'Label Heterogeneity'},
    {"id": '2012.13497', "title": 'Optimal Client Sampling for Federated Learning', "authors": 'Sahu et al.', "year": 2020, "citations": 285, "category": 'Sampling Theory'},
    {"id": '2106.11568', "title": 'Federated Reinforcement Learning for Control', "authors": 'Knott, Duran', "year": 2021, "citations": 95, "category": 'Reinforcement Learning'},
]

LEFTOVER_FILES = ["paper_analysis.xlsx", "literature_review.docx", "distributed_ml_papers.bib"]


def clear_teamly():
    if psycopg2 is None:
        print("[preprocess] psycopg2 unavailable; skip teamly clear")
        return
    try:
        conn = psycopg2.connect(**DB_CONN)
    except Exception as e:
        print(f"[preprocess] teamly DB unavailable, skip: {e}")
        return
    try:
        with conn.cursor() as cur:
            try:
                cur.execute("DELETE FROM teamly.pages WHERE id > 3")
            except Exception:
                pass
        conn.commit()
        print("[preprocess] Cleared user Teamly pages (id>3)")
    finally:
        conn.close()


def write_source_csv(agent_ws: Path):
    """Inject the consumable RU-described source CSV with citation counts."""
    csv_path = agent_ws / "paper_sources.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(SOURCE_HEADER)
        for p in PAPERS:
            w.writerow([p["id"], p["title"], p["authors"], p["year"], p["citations"], p["category"]])
    print(f"[preprocess] Wrote source CSV with {len(PAPERS)} papers -> {csv_path}")


def clear_leftovers(agent_ws: Path):
    for name in LEFTOVER_FILES:
        fp = agent_ws / name
        if fp.exists():
            try:
                fp.unlink()
                print(f"[preprocess] Removed leftover {name}")
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clear_teamly()

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        clear_leftovers(agent_ws)
        write_source_csv(agent_ws)

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
