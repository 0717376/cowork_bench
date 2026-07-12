"""Предобработка: данные moex.* засеяны глобально и доступны только для чтения.

Инъекция/очистка схемы не требуется: moex.stock_prices и moex.stock_info
наполняются глобальным сидом db/zzz_moex_after_init.sql. Файл-ответ
Sector_Analysis.xlsx НЕ создаём (его должен сформировать агент). Идемпотентно
удаляем возможный устаревший Sector_Analysis.xlsx из рабочего каталога агента.
"""
import argparse
import os


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        stale = os.path.join(args.agent_workspace, "Sector_Analysis.xlsx")
        if os.path.exists(stale):
            os.remove(stale)
            print(f"[preprocess] удалён устаревший файл: {stale}")

    print("[preprocess] moex.* доступна только для чтения — инъекция данных не требуется")


if __name__ == "__main__":
    main()
