"""Preprocess: внедрение данных не требуется (источники данных только для чтения)."""
import argparse

def main():
    # Нет схем для очистки — источники данных доступны только для чтения,
    # данные ClickHouse (схема sf_data) сидируются централизованно.
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("No data injection needed - using read-only data")

if __name__ == "__main__":
    main()
