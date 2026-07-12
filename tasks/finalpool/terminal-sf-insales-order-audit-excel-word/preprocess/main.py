"""Preprocess for terminal-sf-insales-order-audit-excel-word.
No writable schemas used. ClickHouse (sf_data.*) and InSales (wc.*) are read-only;
their RU data values are seeded centrally (db/zzz_clickhouse_after_init.sql,
db/zzz_wc_after_init.sql) - no DB injection here.
Just clears any leftover output files from previous runs (idempotent)."""
import argparse
import os
import glob as globmod


def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        for pattern in ["Order_Audit_Report.xlsx", "Audit_Findings.docx", "audit_analysis.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done. No DB injection needed (read-only sources, central RU seeds).")


if __name__ == "__main__":
    main()
