"""Preprocess: внедрение данных не требуется (только read-only источники Canvas)."""
import argparse

def main():
    # Нет записываемых схем для очистки - источники данных только для чтения (Canvas).
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("No data injection needed - using read-only Canvas data")

if __name__ == "__main__":
    main()
