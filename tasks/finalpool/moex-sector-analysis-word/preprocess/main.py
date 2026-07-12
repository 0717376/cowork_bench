"""
Preprocess для yf-sector-analysis-word (RU / moex-finance).
Данные moex.* засеяны глобально (db/zzz_moex_after_init.sql) — записываемых
схем нет, чистить нечего. Источник read-only.
"""

import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("[preprocess] Предобработка не требуется (источник moex.* read-only).")


if __name__ == "__main__":
    main()
