"""Preprocess for moex-peer-comparison-excel-ppt-email. Clears email data and injects RU noise emails."""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        # Дополняем moex.stock_info полями, нужными для peer-comparison:
        # dividendYield (%) и beta задаём только для этой задачи (в общем сиде их нет,
        # т.к. другие moex-задачи документируют их отсутствие), плюс заполняем пустой
        # sector у TCSG. Идемпотентно: jsonb || перезаписывает ключи.
        moex_patch = {
            "SBER.ME": '{"dividendYield": 10.5, "beta": 1.15}',
            "GAZP.ME": '{"dividendYield": 0.0, "beta": 1.05}',
            "LKOH.ME": '{"dividendYield": 7.1, "beta": 0.95}',
            "TCSG.ME": '{"dividendYield": 0.7, "beta": 1.35, "sector": "Financial Services", '
                       '"sectorDisp": "Financial Services", "sectorKey": "financial-services", '
                       '"industry": "Credit Services", "industryDisp": "Credit Services"}',
            "MGNT.ME": '{"dividendYield": 1.85, "beta": 0.70}',
            "MTSS.ME": '{"dividendYield": 12.7, "beta": 0.80}',
        }
        for sym, patch in moex_patch.items():
            cur.execute(
                "UPDATE moex.stock_info SET data = data || %s::jsonb WHERE symbol = %s",
                (patch, sym),
            )

        # Досеиваем годовые cashflow-отчёты (в общем сиде их нет для этих тикеров):
        # FCF соответствует groundtruth Peer_Comparison.xlsx, OCF — operatingCashflow
        # из stock_info, CapEx = FCF - OCF для согласованности. Идемпотентно (ON CONFLICT).
        cashflow_rows = {
            "SBER.ME": (658000000000, 745000000000, -87000000000),
            "GAZP.ME": (-930716647424, 2867018072064, -3797734719488),
            "LKOH.ME": (574309138432, 1717469052928, -1143159914496),
            "TCSG.ME": (-587945016320, -549945016320, -38000000000),
            "MGNT.ME": (-139289952256, 45002563584, -184292515840),
            "MTSS.ME": (202972872704, 274431000576, -71458127872),
        }
        for sym, (fcf, ocf, capex) in cashflow_rows.items():
            cur.execute(
                """
                INSERT INTO moex.financial_statements (symbol, period_end, stmt_type, freq, data)
                VALUES (%s, '2025-12-31', 'cashflow', 'annual',
                        jsonb_build_object('Free Cash Flow', %s::numeric,
                                           'Operating Cash Flow', %s::numeric,
                                           'Capital Expenditure', %s::numeric))
                ON CONFLICT (symbol, period_end, stmt_type, freq)
                DO UPDATE SET data = EXCLUDED.data
                """,
                (sym, fcf, ocf, capex),
            )

        # Clear email tables
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")

        # Get inbox folder id
        cur.execute("SELECT id FROM email.folders WHERE LOWER(name) LIKE '%inbox%' LIMIT 1")
        row = cur.fetchone()
        inbox_fid = row[0] if row else 1

        # Inject noise emails (RU distractors, MOEX-relevant)
        cur.execute("""
            INSERT INTO email.messages (subject, from_addr, to_addr, body_text, folder_id)
            VALUES
            ('Еженедельный обзор рынка - март 2026',
             'markets@newsletter.example.com',
             '["assistant@company.example.com"]',
             'На этой неделе динамика по основным индексам Московской биржи была разнонаправленной. Бумаги нефтегазового сектора продолжили снижение, тогда как акции финансового и потребительского секторов показали относительную силу.',
             %s),
            ('Напоминание: расписание публикаций отчётности за 1 квартал',
             'events@firm.com',
             '["assistant@company.example.com"]',
             'Во вложении расписание публикации финансовой отчётности за 1 квартал 2026 по нашему покрытию. Ключевые даты: SBER 24 апреля, GAZP 22 апреля, LKOH 15 апреля, TCSG 11 апреля, MGNT 25 апреля, MTSS 28 апреля.',
             %s),
            ('Плановое обслуживание ИТ-инфраструктуры офиса',
             'it@firm.com',
             '["assistant@company.example.com"]',
             'Плановые работы в эту субботу с 2:00 до 6:00. Почта и календарь могут быть кратковременно недоступны.',
             %s)
        """, (inbox_fid, inbox_fid, inbox_fid))

        conn.commit()
        print("[preprocess] Cleared email data, injected noise emails, seeded cashflow statements.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()