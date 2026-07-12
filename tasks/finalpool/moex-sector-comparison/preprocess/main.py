"""Preprocess: данные moex.* засеяны глобально и доступны только для чтения.

Инъекция/очистка не требуется: схема moex.stock_info наполняется глобальным
сидом db/zzz_moex_after_init.sql. Файлы-ответы (YF_Sector_Comparison.xlsx и
Sector_Analysis.docx) НЕ создаём — их должен сформировать агент.
"""
import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.parse_args()
    print("[preprocess] moex.* доступна только для чтения — инъекция данных не требуется")


if __name__ == "__main__":
    main()
