import argparse
import os


def main():
    # Источники данных (kulinar MCP, nutrition_guide.md) — только для чтения,
    # схем для очистки в БД нет. Здесь лишь идемпотентно удаляем возможные
    # устаревшие артефакты-ответы из рабочей директории агента, чтобы агент
    # создавал их с нуля. НЕ создаём здесь никаких ответов-заготовок.
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        for fname in ("Nutrition_Comparison.xlsx", "Healthy_Eating_Guide.pptx"):
            path = os.path.join(args.agent_workspace, fname)
            try:
                if os.path.isfile(path):
                    os.remove(path)
                    print(f"Removed stale artifact: {path}")
            except OSError as e:
                print(f"Could not remove {path}: {e}")

    print("Preprocess complete (kulinar recipe task; no DB seed needed)")


if __name__ == "__main__":
    main()
