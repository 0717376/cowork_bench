"""Preprocess for terminal-insales-yf-ppt-notion-email.
Clears teamly (task space) and email. WC (insales) and MOEX are read-only,
globally seeded. Injects RU noise data the agent must ignore. Idempotent."""
import argparse
import json
import os
import uuid
import glob as globmod

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}

# Dedicated Teamly space for this task. The agent's "Market Strategy Tracker"
# page must live here. We clear all pages in this space each run (idempotent).
SPACE_KEY = "MKT_STRATEGY"
SPACE_NAME = "Маркетинговая стратегия"
NOISE_PAGE_TITLE = "Старые заметки по инвентаризации"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # --- Teamly: ensure a dedicated space, clear its pages ---
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES (%s, %s, 'Пространство для стратегического анализа маркетинга и e-commerce.')
                ON CONFLICT (key) DO NOTHING
            """, (SPACE_KEY, SPACE_NAME))
            cur.execute("SELECT id FROM teamly.spaces WHERE key = %s", (SPACE_KEY,))
            space_id = cur.fetchone()[0]
            cur.execute("DELETE FROM teamly.pages WHERE space_id = %s", (space_id,))
            print("[preprocess] Cleared teamly pages in task space.")
        else:
            space_id = None
            print("[preprocess] WARNING: teamly.spaces not found.")

        # Clear email
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

        # --- Inject noise Teamly page (RU) the agent must ignore ---
        if space_id is not None:
            cur.execute(
                "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
                (space_id, NOISE_PAGE_TITLE,
                 "Устаревшая система учёта запасов. Не относится к текущей задаче.",
                 "team"))
            print("[preprocess] Injected noise teamly page.")

        # --- Inject noise email data ---
        cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
        row = cur.fetchone()
        if row:
            inbox_id = row[0]
        else:
            cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX') RETURNING id")
            inbox_id = cur.fetchone()[0]

        noise_emails = [
            ("Заметки с еженедельной планёрки", "admin@company.com", "team@company.com",
             "Заметки с планёрки за эту неделю. Просьба ознакомиться с пунктами действий."),
            ("Обзор бюджета за I квартал", "finance@company.com", "directors@company.com",
             "Во вложении обзор бюджета за I квартал. В целом расходы в пределах целевых значений."),
            ("Подтверждение заказа канцтоваров", "supplies@vendor.com", "office@company.com",
             "Ваш заказ #12345 подтверждён и будет отправлен на следующей неделе."),
        ]
        for subj, from_addr, to_addr, body in noise_emails:
            msg_id = f"<{uuid.uuid4()}@noise.local>"
            cur.execute("""
                INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s, true)
            """, (inbox_id, msg_id, subj, from_addr, json.dumps([to_addr]), body))
        print("[preprocess] Injected noise email data.")

        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Clean any leftover agent outputs
    if args.agent_workspace:
        for pattern in ["Market_Strategy_Presentation.pptx", "market_correlation.py",
                        "category_analysis.py", "market_correlation.json",
                        "category_market_analysis.json"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
