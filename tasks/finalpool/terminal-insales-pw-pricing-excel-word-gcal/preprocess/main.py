"""Preprocess script for terminal-insales-pw-pricing-excel-word-gcal task.
InSales store data (wc.* schema) is read-only and russified centrally.
Clear gcal + inject RU noise. Setup mock competitor pricing server on port 30417
(russified category headers, identical prices / 19 rows).
"""
import argparse
import asyncio
import glob as globmod
import os
import shutil
import tarfile
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


def clear_gcal(cur):
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] Google Calendar events cleared.")
    # Inject gcal noise
    cur.execute("INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status) VALUES (%s, %s, %s, %s, %s, 'confirmed')",
        (str(uuid.uuid4()), "Ежедневный стендап", "Регулярная планёрка команды", "2026-03-05 09:00:00", "2026-03-05 09:30:00"))
    cur.execute("INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status) VALUES (%s, %s, %s, %s, %s, 'confirmed')",
        (str(uuid.uuid4()), "Йога в обед", "Оздоровительное мероприятие", "2026-03-06 12:00:00", "2026-03-06 12:45:00"))
    print("[preprocess] Injected noise data.")


async def setup_mock_server(port=30417):
    print("[preprocess] Setting up mock competitor pricing website...")
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")

    mock_dir = os.path.join(tmp_dir, "mock_pages")

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    # Launch fully detached (own session via start_new_session) and bind explicitly
    # to 0.0.0.0 so the server survives this asyncio loop tearing down and is
    # reachable over IPv4 localhost / the container IP from the in-container
    # Playwright (stdio) browser. Previously the bare `nohup ... &` could be reaped
    # on loop close and bound only ::1, leaving http://localhost:30417 unreachable.
    log_path = os.path.join(mock_dir, "server.log")
    proc = await asyncio.create_subprocess_exec(
        "python3", "-m", "http.server", str(port),
        "--bind", "0.0.0.0", "--directory", mock_dir,
        stdout=open(log_path, "wb"),
        stderr=asyncio.subprocess.STDOUT,
        stdin=asyncio.subprocess.DEVNULL,
        start_new_session=True,
    )
    print(f"[preprocess] Mock server pid={proc.pid} on 0.0.0.0:{port}")

    # Readiness probe: poll the served page over IPv4 localhost until it answers.
    ready = False
    for _ in range(30):
        await asyncio.sleep(0.25)
        check = await asyncio.create_subprocess_shell(
            f"python3 -c \"import urllib.request,sys; "
            f"sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:{port}/pricing.html', timeout=2).status==200 else 1)\"",
            stdout=asyncio.subprocess.DEVNULL, stderr=asyncio.subprocess.DEVNULL,
        )
        await check.wait()
        if check.returncode == 0:
            ready = True
            break
    if ready:
        print(f"[preprocess] Mock competitor website ready at http://localhost:{port}/pricing.html")
    else:
        print(f"[preprocess] WARNING: mock server on port {port} did not become ready")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gcal(cur)
        conn.commit()
        print("[preprocess] DB operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server(30417)

    if args.agent_workspace:
        for pattern in ["Competitive_Pricing_Report.xlsx", "Pricing_Strategy_Report.docx", "price_analysis_output.txt"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
