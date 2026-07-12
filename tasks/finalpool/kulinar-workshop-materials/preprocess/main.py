"""
Preprocess для задачи kulinar-workshop-materials.

Эта задача не использует записываемые схемы БД.
Агент получает рецепты из MCP-сервера kulinar во время выполнения
(база kulinar засеяна глобально, отдельная инициализация БД не требуется).
Preprocess только гарантирует наличие рабочей директории и идемпотентно
удаляет возможные оставшиеся выходные файлы предыдущих запусков.
"""
import argparse
import os


# Имена выходных файлов (английские — eval ищет их по точному имени).
OUTPUT_FILES = [
    "Workshop_Handbook.docx",
    "Workshop_Slides.pptx",
    "Shopping_List.pdf",
]


def main():
    # Записываемых схем для DELETE нет — источники данных только для чтения.
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False, help="Launch time")
    args = parser.parse_args()

    if args.agent_workspace:
        os.makedirs(args.agent_workspace, exist_ok=True)
        print(f"Рабочая директория агента готова: {args.agent_workspace}")

        # Идемпотентная очистка оставшихся выходных файлов (без пресидинга ответа).
        for name in OUTPUT_FILES:
            path = os.path.join(args.agent_workspace, name)
            if os.path.exists(path):
                try:
                    os.remove(path)
                    print(f"Удалён оставшийся файл: {path}")
                except OSError as e:
                    print(f"Не удалось удалить {path}: {e}")

    print("Preprocess завершён — инициализация БД для этой задачи не нужна.")


if __name__ == "__main__":
    main()
