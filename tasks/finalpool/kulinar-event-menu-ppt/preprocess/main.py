"""
Preprocess для kulinar-event-menu-ppt (русифицированная версия).

База рецептов kulinar — standalone MCP-сервер (данные загружаются из
all_recipes.json), отдельная схема в БД не нужна. Здесь только чистим
рабочую директорию от прошлых результатов для идемпотентности.
НЕ создаём заранее .pptx/.xlsx — иначе задача автоматически зачтётся.
"""

import argparse
import os
import glob


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    workspace = args.agent_workspace
    if workspace:
        # Удаляем возможные результаты прошлого запуска
        for pattern in ["Event_Menu_Presentation.pptx", "Menu_Budget.xlsx"]:
            for f in glob.glob(os.path.join(workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Удалён {f}")

    print("[preprocess] Готово.")


if __name__ == "__main__":
    main()
