"""Предобработка: инъекция данных не требуется (только источники, доступные на чтение)."""
import argparse

def main():
    # Нет записываемых схем для очистки — источники данных доступны только на чтение
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("Инъекция данных не требуется — используются данные только для чтения")

if __name__ == "__main__":
    main()
