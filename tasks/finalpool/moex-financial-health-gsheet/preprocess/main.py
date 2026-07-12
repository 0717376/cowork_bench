"""
Preprocess for moex-financial-health-gsheet task.
- Clears gsheet schema tables
- Blanks Operating Income / Gross Profit for the bank ticker (SBER.ME) so the
  N/A code path stays exercised (banks may not report these lines)
- Extracts mock_pages.tar.gz and starts HTTP server on port 30158
"""
import argparse
import asyncio
import os
import shutil
import tarfile
import psycopg2


DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

PORT = 30158


def setup_gsheet():
    """Clear all gsheet tables so the agent starts fresh."""
    conn = psycopg2.connect(**DB_CONN)
    conn.autocommit = True
    cur = conn.cursor()

    print("Clearing gsheet tables ...")
    cur.execute("DELETE FROM gsheet.cells;")
    cur.execute("DELETE FROM gsheet.permissions;")
    cur.execute("DELETE FROM gsheet.sheets;")
    cur.execute("DELETE FROM gsheet.spreadsheets;")
    cur.execute("DELETE FROM gsheet.folders;")
    print("  -> All gsheet tables cleared.")

    cur.close()
    conn.close()


# Bank ticker that should NOT report Gross Profit / Operating Income, so the
# agent must record Operating_Margin_Pct = N/A for it. The moex seed populates
# these lines for every ticker, so we blank them here (idempotent).
BANK_SYMBOL = "SBER.ME"
BLANK_KEYS = ["Operating Income", "Gross Profit"]


def setup_bank_na_path():
    """Remove Operating Income / Gross Profit from the bank ticker income
    statements so the N/A pedagogy is reproduced. Idempotent."""
    conn = psycopg2.connect(**DB_CONN)
    conn.autocommit = True
    cur = conn.cursor()

    print(f"Blanking {BLANK_KEYS} for {BANK_SYMBOL} income statements ...")
    for key in BLANK_KEYS:
        cur.execute(
            """
            UPDATE moex.financial_statements
            SET data = data - %s
            WHERE symbol = %s AND stmt_type = 'income_stmt'
            """,
            (key, BANK_SYMBOL),
        )
    print(f"  -> {BANK_SYMBOL} now lacks Gross Profit / Operating Income.")

    cur.close()
    conn.close()


async def run_command(cmd: str):
    proc = await asyncio.create_subprocess_shell(
        cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE
    )
    await proc.wait()


async def setup_mock_server():
    """Extract mock pages and start HTTP server."""
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")

    print("Setting up mock credit analysis portal ...")
    tmp_dir = os.path.join(task_root, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"  -> Extracted {tar_path} to {tmp_dir}")

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    await run_command(f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null")
    await asyncio.sleep(0.5)
    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {PORT} --directory \"{mock_dir}\" "
        f"> \"{mock_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"  -> Mock credit analysis portal running at http://localhost:{PORT}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    # 1. Clear Google Sheet data in database
    setup_gsheet()

    # 2. Blank bank ticker's Gross Profit / Operating Income (N/A path)
    setup_bank_na_path()

    # 3. Set up mock HTTP server
    await setup_mock_server()

    print("\nPreprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
