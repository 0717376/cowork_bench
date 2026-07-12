"""Preprocess: инъекция данных не требуется.

Схема moex.* (источник данных moex-finance) уже глобально засеяна и доступна
только для чтения. Агент получает котировки через MCP moex-finance, поэтому
здесь нечего удалять или подготавливать.
"""
import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("Инъекция данных не требуется — используются read-only данные moex.*")


if __name__ == "__main__":
    main()
