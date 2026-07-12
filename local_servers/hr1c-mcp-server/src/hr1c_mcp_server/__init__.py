"""hr1c MCP-сервер (1С:ЗУП-подобный HR-сервер поверх PostgreSQL).

Предоставляет инструменты list_tables / describe_table / read_query / write_query
с 3-частной нотацией БАЗА.СХЕМА.ТАБЛИЦА (как Snowflake/1C). Под капотом — PostgreSQL,
схема hr1c_data, таблицы вида "БАЗА__СХЕМА__ТАБЛИЦА".
"""
import argparse
import asyncio
import logging
import os
import sys

import dotenv

from . import server


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--allow_write", required=False, default=False, action="store_true",
                        help="Разрешить операции записи в БД")
    parser.add_argument("--log_dir", required=False, default=None, help="Каталог для логов")
    parser.add_argument("--log_level", required=False, default="INFO", help="Уровень логирования")
    parser.add_argument("--prefetch", action="store_true", dest="prefetch", default=False,
                        help="Предзагрузить описания таблиц")
    parser.add_argument("--no-prefetch", action="store_false", dest="prefetch")
    parser.add_argument("--exclude_tools", required=False, default=[], nargs="+",
                        help="Список инструментов для исключения")
    parser.add_argument("--exclude-json-results", action="store_true",
                        dest="exclude_json_results", default=False,
                        help="Не возвращать JSON-результаты")
    parser.add_argument("--allowed_databases", required=False,
                        help="Список БАЗ через запятую, к которым разрешён доступ")
    # Эти аргументы оставлены ради совместимости с CLI-схемой Snowflake-style серверов
    # (private_key_path и т.п.). Игнорируются: в hr1c аутентификация по PG_USER/PG_PASSWORD.
    parser.add_argument("--private_key_path", required=False, help="игнорируется")

    args, unknown = parser.parse_known_args()

    connection_args = {}
    for i in range(0, len(unknown), 2):
        if i + 1 >= len(unknown):
            break
        key = unknown[i]
        value = unknown[i + 1]
        if key.startswith("--"):
            key = key[2:]
            connection_args[key] = value

    allowed_databases = None
    if args.allowed_databases:
        allowed_databases = [db.strip() for db in args.allowed_databases.split(',')]
        logging.warning(f"Allowed databases: {allowed_databases}")

    server_args = {
        "allow_write": args.allow_write,
        "log_dir": args.log_dir,
        "log_level": args.log_level,
        "prefetch": args.prefetch,
        "exclude_tools": args.exclude_tools,
        "exclude_json_results": args.exclude_json_results,
        "allowed_databases": allowed_databases,
    }
    return server_args, connection_args


def main():
    dotenv.load_dotenv()
    server_args, connection_args = parse_args()

    # database/schema нужны для совместимости с интерфейсом (никаких реальных
    # snowflake-объектов не создаём; реальное подключение — через PG_* env).
    connection_args.setdefault("database", os.environ.get("PG_DATABASE", "cowork_gym"))
    connection_args.setdefault("schema", "hr1c")

    asyncio.run(
        server.main(
            connection_args=connection_args,
            allow_write=server_args["allow_write"],
            log_dir=server_args["log_dir"],
            prefetch=server_args["prefetch"],
            log_level=server_args["log_level"],
            exclude_tools=server_args["exclude_tools"],
            exclude_json_results=server_args["exclude_json_results"],
            allowed_databases=server_args["allowed_databases"],
        )
    )


__all__ = ["main", "server", "write_detector"]

if __name__ == "__main__":
    main()
