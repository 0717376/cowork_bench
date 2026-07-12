#!/usr/bin/env bash
# Smoke test для db-миграций.
#
# Поднимает свежий postgres:15 со ВСЕМИ файлами из db/*.sql* в
# /docker-entrypoint-initdb.d/, ждёт healthy ≤60s. Если init упал — печатает
# первую ERROR/FATAL строку и выходит с кодом 1.
#
# Запускать перед коммитом любого изменения в db/ или после переименования
# init-файлов. Ловит классику: новый файл сортируется до оригинального dump
# по en_US.utf8 collation и фейлит entrypoint (см. docs/ENVIRONMENT_QUIRKS.md).
#
# Usage:  scripts/test_db_migration.sh
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$REPO_ROOT"

CONTAINER="pg-migrate-smoke-$$"
TIMEOUT_S=120

cleanup() {
    docker rm -fv "$CONTAINER" >/dev/null 2>&1 || true
}
trap cleanup EXIT

mounts=()
for f in db/*.sql db/*.sql.gz; do
    [ -e "$f" ] || continue
    mounts+=(-v "$(pwd)/${f}:/docker-entrypoint-initdb.d/$(basename "$f"):ro")
done

if [ ${#mounts[@]} -eq 0 ]; then
    echo "[smoke] no SQL files found in db/" >&2
    exit 2
fi

echo "[smoke] mounting ${#mounts[@]} file(s) into a fresh postgres:15"
docker rm -fv "$CONTAINER" >/dev/null 2>&1 || true
docker run -d --name "$CONTAINER" \
    -e POSTGRES_DB=cowork_gym \
    -e POSTGRES_USER=eigent \
    -e POSTGRES_PASSWORD=camel \
    "${mounts[@]}" \
    --health-cmd="pg_isready -U eigent -d cowork_gym" \
    --health-interval=2s --health-retries=60 \
    postgres:15 >/dev/null

start=$(date +%s)
while :; do
    elapsed=$(($(date +%s) - start))
    status=$(docker inspect --format '{{.State.Health.Status}}' "$CONTAINER" 2>/dev/null || echo "missing")

    if [ "$status" = "healthy" ]; then
        # entrypoint может пометить healthy раньше, чем все init-файлы дочитаны;
        # подождём пока в логах появится финальное "ready for start up".
        for i in $(seq 1 30); do
            if docker logs "$CONTAINER" 2>&1 | grep -q "PostgreSQL init process complete; ready for start up"; then
                break
            fi
            sleep 1
        done
        elapsed=$(($(date +%s) - start))
        echo "[smoke] OK — init complete at ${elapsed}s"
        echo "[smoke] init order observed:"
        docker logs "$CONTAINER" 2>&1 | grep "docker-entrypoint.sh: running" \
            | sed 's|^|    |'
        # sanity-проверка: убедимся, что ни один файл не упал тихо.
        if docker logs "$CONTAINER" 2>&1 | grep -qE '^psql:.*ERROR|^.*FATAL'; then
            echo "[smoke] FAIL — init produced errors:" >&2
            docker logs "$CONTAINER" 2>&1 | grep -E '^psql:.*ERROR|^.*FATAL' | head -5 >&2
            exit 1
        fi
        exit 0
    fi

    if [ "$elapsed" -ge "$TIMEOUT_S" ]; then
        echo "[smoke] FAIL — postgres not healthy after ${TIMEOUT_S}s" >&2
        echo "[smoke] first ERROR/FATAL line:" >&2
        docker logs "$CONTAINER" 2>&1 | grep -m1 -E '(ERROR|FATAL)' >&2 || true
        echo "[smoke] init scripts that ran:" >&2
        docker logs "$CONTAINER" 2>&1 | grep "docker-entrypoint.sh: running" \
            | sed 's|^|    |' >&2 || true
        exit 1
    fi
    sleep 2
done
