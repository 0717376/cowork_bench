"""Preprocess for terminal-moex-fetch-gsheet-excel-teamly.
MOEX finance is read-only. Set up mock HTTP server, clear gsheet, idempotently
remove leftover "Investment Research Log" teamly pages, inject RU noise.
We do NOT pre-create the deliverable pages/spreadsheet (those are the agent's job).
The seeded teamly spaces (TEAM/TRIPS) and their seed pages are preserved."""
import argparse
import glob
import os
import shutil
import subprocess
import tarfile
import time
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORT = 30180


def get_conn():
    return psycopg2.connect(**DB_CONFIG)


def clear_writable(cur):
    print("[preprocess] Clearing gsheet data; removing leftover teamly deliverables...")
    # Google Sheets (agent recreates the deliverable spreadsheet)
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.sheets")
    cur.execute("DELETE FROM gsheet.permissions")
    cur.execute("DELETE FROM gsheet.spreadsheets")

    # Teamly: idempotently remove only leftover deliverable pages from prior runs.
    # Seeded TEAM/TRIPS spaces and their seed pages are preserved.
    cur.execute("DELETE FROM teamly.pages WHERE lower(title) LIKE '%investment research log%'")


def inject_noise(cur):
    print("[preprocess] Injecting noise data...")
    # Noise teamly pages in the existing TEAM space (RU titles, plausible content).
    cur.execute("SELECT id FROM teamly.spaces WHERE key = 'TEAM'")
    row = cur.fetchone()
    if row:
        team_id = row[0]
        noise_pages = [
            ("Архив протоколов совещаний",
             "Сводный архив протоколов еженедельных совещаний аналитического отдела."),
            ("Планирование Q4",
             "Черновик целей и приоритетов отдела на четвёртый квартал."),
            ("Ретроспектива команды",
             "Заметки по итогам спринта: что прошло хорошо, что улучшить."),
        ]
        for title, body in noise_pages:
            # Idempotent: skip if a noise page with this title already exists.
            cur.execute(
                "SELECT 1 FROM teamly.pages WHERE space_id = %s AND title = %s",
                (team_id, title),
            )
            if cur.fetchone() is None:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) "
                    "VALUES (%s, %s, %s, %s)",
                    (team_id, title, body, "Аналитический отдел"),
                )

    # Noise Google Sheet
    noise_ss_id = str(uuid.uuid4())
    cur.execute("""
        INSERT INTO gsheet.spreadsheets (id, title)
        VALUES (%s, 'Бюджет Q4 (черновик)')
    """, (noise_ss_id,))
    cur.execute("""
        INSERT INTO gsheet.sheets (spreadsheet_id, title, "index", row_count, column_count)
        VALUES (%s, 'Sheet1', 0, 100, 10)
    """, (noise_ss_id,))


def setup_mock_server():
    files_dir = os.path.join(TASK_ROOT, "files")
    # TASK_ROOT/tmp lives on the read-only tasks mount in the agent container,
    # so rmtree/makedirs there raise OSError and crash preprocess (the mock
    # feed never starts). Serve from a guaranteed-writable path instead.
    tmp_dir = os.environ.get("MOCK_SERVE_DIR", f"/tmp/mock_pages_{PORT}")

    # Kill existing process on port. lsof may be absent in the agent image;
    # fall back to fuser/ss so we never crash on a missing binary.
    for kill_cmd in (
        f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null",
        f"fuser -k {PORT}/tcp 2>/dev/null",
    ):
        try:
            subprocess.run(kill_cmd, shell=True, timeout=5)
        except Exception:
            pass
    time.sleep(0.5)

    # Guard rmtree/makedirs against a read-only filesystem.
    try:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
        os.makedirs(tmp_dir, exist_ok=True)
    except OSError as e:
        print(f"[preprocess] WARN: cannot reset {tmp_dir} ({e}); using fallback")
        tmp_dir = f"/tmp/mock_pages_{PORT}_{uuid.uuid4().hex[:8]}"
        os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    if os.path.exists(mock_dir):
        log_path = os.path.join(mock_dir, "server.log")
        subprocess.Popen(
            f"nohup python3 -m http.server {PORT} --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"[preprocess] Mock server started on port {PORT}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    conn = get_conn()
    cur = conn.cursor()
    try:
        clear_writable(cur)
        inject_noise(cur)
        conn.commit()
        print("[preprocess] DB setup done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    setup_mock_server()

    if args.agent_workspace:
        for pattern in ["Market_Analysis_Report.xlsx", "market_*.py", "market_*.json",
                        "stock_*.json", "economic_*.json"]:
            for f in glob.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
