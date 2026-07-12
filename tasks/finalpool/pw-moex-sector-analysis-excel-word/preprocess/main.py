"""Preprocess для pw-moex-sector-analysis-excel-word (RU / moex-finance).

Данные moex.* засеяны глобально (db/zzz_moex_after_init.sql) и доступны
только для чтения — никакой инъекции/очистки PG-схем не требуется.

Что делаем:
  1) Идемпотентно убираем возможные «остатки» в рабочей директории агента
     (Sector_Analysis_Report.xlsx, Sector_Analysis_Analysis.docx,
     yf_sector_results.json, yf_sector_processor.py и временные JSON), чтобы
     агент строил всё с нуля. ЭТАЛОН НЕ инъектируется.
  2) Поднимаем локальный mock-сервер с эталонной отраслевой панелью
     (http://localhost:30315), который агент читает через playwright.
"""
import argparse
import os
import shutil
import subprocess
import tarfile
import time

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Файлы-ответы агента: их преднамеренно НЕ создаём, только подчищаем остатки.
LEFTOVER_FILES = [
    "Sector_Analysis_Report.xlsx",
    "Sector_Analysis_Analysis.docx",
    "yf_sector_results.json",
    "yf_sector_processor.py",
]


def clear_agent_workspace(agent_workspace):
    """Идемпотентно удаляем возможные прошлые артефакты агента."""
    if not agent_workspace or not os.path.isdir(agent_workspace):
        return
    for name in LEFTOVER_FILES:
        path = os.path.join(agent_workspace, name)
        try:
            if os.path.isfile(path):
                os.remove(path)
                print(f"[preprocess] удалён остаток: {path}")
        except OSError as e:
            print(f"[preprocess] не удалось удалить {path}: {e}")


def setup_mock_server(port=30315):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Освобождаем порт от предыдущего процесса
    try:
        subprocess.run(f"kill -9 $(lsof -ti:{port}) 2>/dev/null", shell=True, timeout=5)
    except Exception:
        pass
    time.sleep(0.5)

    # Распаковываем mock-страницы
    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)

    # Запускаем HTTP-сервер
    mock_dir = os.path.join(tmp_dir, "mock_pages")
    if os.path.exists(mock_dir):
        log_path = os.path.join(mock_dir, "server.log")
        subprocess.Popen(
            f"nohup python3 -m http.server {port} --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &",
            shell=True,
        )
        time.sleep(1)
        print(f"Mock server started on port {port}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_agent_workspace(args.agent_workspace)
    setup_mock_server(30315)


if __name__ == "__main__":
    main()
