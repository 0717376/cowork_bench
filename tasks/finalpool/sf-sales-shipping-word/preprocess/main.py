"""Preprocess: данные не инжектируются (источник — read-only база ClickHouse)."""
import argparse

def main():
    # Нет записываемых схем для очистки — источник данных только для чтения.
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("[preprocess] Инжекция данных не требуется — используются read-only данные ClickHouse (sf_data).")
    print("[preprocess] Done.")

if __name__ == "__main__":
    main()
