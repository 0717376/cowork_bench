"""
Preprocess script for kulinar-scholarly-health-study task.

This script:
1. Clears scholarly data and injects nutrition-related papers
2. Extracts mock_pages.tar.gz and starts HTTP server on port 30155
"""

import argparse
import asyncio
import json
import os
import shutil
import tarfile

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

# Injected diet/health papers (russified). arxiv_id format, q-bio.OT category,
# and integer citation_count preserved for the eval substring/value checks.
PAPERS = [
    {
        "arxiv_id": "2001.00001",
        "title": "Структура рациона питания и риск хронических заболеваний",
        "authors": [
            {"name": "Светлана Иванова (Svetlana Ivanova)"},
            {"name": "Михаил Орлов (Mikhail Orlov)"},
            {"name": "Лариса Ванг (Larisa Wang)"},
            {"name": "Дмитрий Ким (Dmitry Kim)"},
        ],
        "categories": ["q-bio.OT"],
        "primary_category": "q-bio.OT",
        "abstract": (
            "В этом обзоре исследуется связь между структурой рациона питания и риском "
            "хронических заболеваний, включая сердечно-сосудистые заболевания, диабет и "
            "онкологические заболевания. Проанализированы данные 45 проспективных когортных "
            "исследований в разных странах и кулинарных традициях. Результаты показывают, "
            "что рацион, богатый цельнозерновыми продуктами, овощами и нежирным белком, "
            "связан со значительным снижением риска хронических заболеваний."
        ),
        "published": "2020-01-15",
        "pub_year": 2020,
        "venue": "Вопросы питания",
        "citation_count": 500,
    },
    {
        "arxiv_id": "2002.00002",
        "title": "Традиционная русская кухня и сердечно-сосудистое здоровье",
        "authors": [
            {"name": "Мария Гаврилова (Maria Gavrilova)"},
            {"name": "Антон Романов (Anton Romanov)"},
            {"name": "Елена Костина (Elena Kostina)"},
        ],
        "categories": ["q-bio.OT"],
        "primary_category": "q-bio.OT",
        "abstract": (
            "Традиционный русский рацион, для которого характерно потребление круп, "
            "корнеплодов, рыбы, ферментированных овощей и умеренного количества жиров, "
            "подробно изучается с точки зрения его влияния на сердечно-сосудистое здоровье. "
            "Данный метаанализ 23 рандомизированных контролируемых исследований "
            "демонстрирует значимое снижение артериального давления и маркеров воспаления "
            "у приверженцев сбалансированного традиционного рациона."
        ),
        "published": "2020-06-20",
        "pub_year": 2020,
        "venue": "Кардиологический вестник",
        "citation_count": 800,
    },
    {
        "arxiv_id": "2003.00003",
        "title": "Растительные рационы питания: обзор доказательной базы",
        "authors": [
            {"name": "Яков Чернов (Yakov Chernov)"},
            {"name": "Раиса Зеленова (Raisa Zelenova)"},
            {"name": "Тимофей Лысенко (Timofey Lysenko)"},
        ],
        "categories": ["q-bio.OT"],
        "primary_category": "q-bio.OT",
        "abstract": (
            "Растительные рационы питания приобрели значительную популярность в последние "
            "годы. В этом обзоре анализируются доказательства пользы для здоровья "
            "растительных моделей питания, включая вегетарианские рационы. Рассматриваются "
            "нутриентная полноценность, влияние на массу тела, факторы сердечно-сосудистого "
            "риска и профилактику заболеваний. Доказательства подтверждают, что грамотно "
            "составленные растительные рационы полноценны и полезны для профилактики "
            "заболеваний."
        ),
        "published": "2021-03-10",
        "pub_year": 2021,
        "venue": "Нутриенты",
        "citation_count": 350,
    },
    {
        "arxiv_id": "2004.00004",
        "title": "Калорийная плотность блюд и контроль массы тела",
        "authors": [
            {"name": "Алексей Жданов (Alexey Zhdanov)"},
            {"name": "Евгения Ушакова (Evgenia Ushakova)"},
        ],
        "categories": ["q-bio.OT"],
        "primary_category": "q-bio.OT",
        "abstract": (
            "В работе исследуется связь между калорийной плотностью повседневных блюд и "
            "долгосрочным контролем массы тела. На основе данных популяционного "
            "исследования показано, что блюда с высокой калорийной плотностью связаны с "
            "повышенным потреблением энергии, тогда как блюда на основе овощей и круп "
            "способствуют поддержанию здорового веса. Обсуждаются практические "
            "рекомендации по составлению рациона."
        ),
        "published": "2021-07-01",
        "pub_year": 2021,
        "venue": "Профилактическая медицина",
        "citation_count": 420,
    },
]


def clear_scholarly(cur):
    """Clear scholarly data."""
    print("[preprocess] Clearing scholarly data...")
    cur.execute("DELETE FROM scholarly.arxiv_papers")
    cur.execute("DELETE FROM scholarly.scholar_papers")
    print("[preprocess] Scholarly data cleared.")


def inject_scholarly_arxiv(cur):
    """Inject papers into scholarly.arxiv_papers."""
    for p in PAPERS:
        cur.execute("""
            INSERT INTO scholarly.arxiv_papers
            (id, title, authors, abstract, categories, primary_category,
             published, updated, doi, journal_ref, pdf_url, html_url, comment)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (id) DO UPDATE SET
                title = EXCLUDED.title,
                authors = EXCLUDED.authors,
                abstract = EXCLUDED.abstract,
                categories = EXCLUDED.categories,
                primary_category = EXCLUDED.primary_category,
                published = EXCLUDED.published
        """, (
            p["arxiv_id"],
            p["title"],
            json.dumps(p["authors"]),
            p["abstract"],
            json.dumps(p["categories"]),
            p["primary_category"],
            p["published"],
            p["published"],
            None,
            p.get("venue"),
            f"http://arxiv.org/pdf/{p['arxiv_id']}",
            f"http://arxiv.org/abs/{p['arxiv_id']}",
            None,
        ))
    print(f"[preprocess] Injected {len(PAPERS)} papers into scholarly.arxiv_papers")


def inject_scholarly_scholar(cur):
    """Inject papers into scholarly.scholar_papers."""
    for p in PAPERS:
        cur.execute("""
            INSERT INTO scholarly.scholar_papers
            (title, authors, abstract, pub_year, venue, citation_count,
             url, eprint_url, pub_url, bib)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            p["title"],
            json.dumps(p["authors"]),
            p["abstract"],
            p["pub_year"],
            p.get("venue"),
            p.get("citation_count", 0),
            f"http://arxiv.org/abs/{p['arxiv_id']}",
            f"http://arxiv.org/pdf/{p['arxiv_id']}",
            f"http://arxiv.org/abs/{p['arxiv_id']}",
            json.dumps({"title": p["title"], "year": p["pub_year"]}),
        ))
    print(f"[preprocess] Injected {len(PAPERS)} papers into scholarly.scholar_papers")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30155."""
    print("[preprocess] Setting up mock nutrition data server...")

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    port = 30155

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --directory \"{mock_dir}\" "
        f"> \"{mock_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock nutrition server running at http://localhost:{port}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_scholarly(cur)
        inject_scholarly_arxiv(cur)
        inject_scholarly_scholar(cur)
        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
