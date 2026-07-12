"""Preprocess: данные moex.* засеяны глобально и доступны только для чтения.

Никакой инъекции/очистки не требуется: схема moex.financial_statements
наполняется глобальным сидом db/zzz_moex_after_init.sql. Файл-ответ
YF_Financial_Statements.xlsx НЕ создаём (его должен сформировать агент).
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
