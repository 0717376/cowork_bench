"""
Скрипт предобработки для задачи kulinar-ecommerce-bundle.

InSales (wc.*), MOEX Finance (moex.*) и Кулинар — все они являются
источниками данных только для чтения и засеваются глобально.
В этой задаче нет записываемых схем, поэтому предобработка минимальна.
"""

import argparse
import os


def main():
    # Нет записываемых схем для очистки — источники данных только для чтения.
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    workspace = args.agent_workspace
    os.makedirs(workspace, exist_ok=True)

    print("[preprocess] Нет записываемых схем для инъекции. Рабочая директория готова.")
    print(f"[preprocess] Рабочая директория агента: {workspace}")
    print("[preprocess] Готово.")


if __name__ == "__main__":
    main()
