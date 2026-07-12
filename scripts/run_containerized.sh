#!/bin/bash
# Run a single task in an ephemeral container with per-task filesystem isolation.
#
# Isolation strategy (aligned with Toolathlon):
#   - Each task runs inside a fresh Docker container that is destroyed on exit.
#   - The container has its own filesystem and MCP server processes.
#   - Postgres (cowork_pg) is shared across tasks; tasks must run sequentially
#     because preprocess resets shared schema state.
#   - A lock file (./dumps/.run.lock) enforces sequential execution.
#
# Prerequisites:
#   1. Build the image:    docker build -t cowork-pack:latest .
#   2. Start postgres:     docker compose up -d postgres
#
# Usage:
#   bash scripts/run_containerized.sh <task_name> [max_steps] [image]
#
# Model configuration (environment variables):
#   MODEL_PROVIDER   Provider key used by main.py: aihubmix | openai | anthropic |
#                    gemini | deepseek | openai_compatible  (overrides eval_config.json)
#   MODEL_NAME       Model name, e.g. gpt-4o, claude-3-5-sonnet-20241022
#   MODEL_API_KEY    API key for the selected provider
#   MODEL_API_URL    Base URL (required for openai_compatible / aihubmix)
#
# Examples:
#   # Native OpenAI:
#   MODEL_PROVIDER=openai MODEL_NAME=gpt-4o \
#     MODEL_API_KEY=sk-proj-xxx \
#     bash scripts/run_containerized.sh insales-coupon-campaign-gcal-gform
#
#   # Via aihubmix (OpenAI-compatible endpoint):
#   MODEL_PROVIDER=aihubmix MODEL_NAME=claude-3-5-sonnet-20241022 \
#     MODEL_API_KEY=sk-xxx \
#     bash scripts/run_containerized.sh kulinar-meal-plan-gcal
#
#   # Native Anthropic:
#   MODEL_PROVIDER=anthropic MODEL_NAME=claude-3-5-haiku-20241022 \
#     MODEL_API_KEY=sk-ant-xxx \
#     bash scripts/run_containerized.sh kulinar-meal-plan-gcal 50

set -euo pipefail

# ---------------------------------------------------------------------------
# Arguments
# ---------------------------------------------------------------------------
TASK="${1:?Usage: $0 <task_name> [max_steps] [image]}"
MAX_STEPS="${2:-100}"
# Image priority: positional $3 > env IMAGE > default. Lets run_parallel.sh or
# the user export IMAGE once instead of repeating it for every invocation.
IMAGE="${3:-${IMAGE:-cowork-pack:latest}}"

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
TASK_SOURCE="$PROJECT_ROOT/tasks/finalpool/$TASK"
DUMPS_DIR="$PROJECT_ROOT/dumps"
LOCK_FILE="$DUMPS_DIR/.run.lock"

# ---------------------------------------------------------------------------
# Validate task directory exists in the image source tree
# ---------------------------------------------------------------------------
if [[ ! -d "$TASK_SOURCE" ]]; then
    echo "[error] Task directory not found: $TASK_SOURCE" >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# Container naming
# ---------------------------------------------------------------------------
TIMESTAMP="$(date +%Y%m%d-%H%M%S)"
SAFE_TASK="$(echo "$TASK" | tr '/' '-')"
# Two ephemeral containers per run (groundtruth isolation):
#   AGENT_CONTAINER — runs the LLM agent; mounts a SANITIZED task tree WITHOUT
#                     groundtruth_workspace/ or evaluation/ (the answer key).
#   EVAL_CONTAINER  — runs the deterministic grader; mounts the FULL task tree
#                     (with groundtruth) but no agent, against the persisted output.
AGENT_CONTAINER="cowork-${SAFE_TASK}-${TIMESTAMP}-agent"
EVAL_CONTAINER="cowork-${SAFE_TASK}-${TIMESTAMP}-eval"

# Output on the host: dumps/<task>/<timestamp>/
# Mounted into the container as /workspace/dumps so that eval_config's
# dump_path="./dumps/" resolves to /workspace/dumps/ inside the container.
OUTPUT_DIR="$DUMPS_DIR/$TASK/$TIMESTAMP"
mkdir -p "$OUTPUT_DIR" "$DUMPS_DIR"

# Sanitized task tree for the agent phase: a copy of ONLY the current task,
# with groundtruth_workspace/ and evaluation/ stripped. Mounted at
# /workspace/tasks it fully shadows the image-baked tasks dir (incl. every
# other task's groundtruth), so the agent cannot read any answer key.
STAGING_DIR="$(mktemp -d "${TMPDIR:-/tmp}/toolathlon-stage-${SAFE_TASK}.XXXXXX")"

# ---------------------------------------------------------------------------
# Logging helpers
# ---------------------------------------------------------------------------
log()  { echo "[$(date +%H:%M:%S)] $*"; }
warn() { echo "[$(date +%H:%M:%S)] [warn] $*" >&2; }
die()  { echo "[$(date +%H:%M:%S)] [error] $*" >&2; exit 1; }

# ---------------------------------------------------------------------------
# Cleanup: stop and remove the ephemeral container on any exit
# ---------------------------------------------------------------------------
cleanup() {
    log "Cleaning up containers ..."
    docker stop "$AGENT_CONTAINER" "$EVAL_CONTAINER" >/dev/null 2>&1 || true
    docker rm   "$AGENT_CONTAINER" "$EVAL_CONTAINER" >/dev/null 2>&1 || true
    [[ -n "${STAGING_DIR:-}" && -d "$STAGING_DIR" ]] && rm -rf "$STAGING_DIR"
    log "Containers removed."
}
trap cleanup EXIT

# ---------------------------------------------------------------------------
# Prerequisites check
# ---------------------------------------------------------------------------
check_prerequisites() {
    command -v docker >/dev/null 2>&1 || die "docker not found in PATH"

    # Verify the image exists locally
    if ! docker image inspect "$IMAGE" >/dev/null 2>&1; then
        die "Image '$IMAGE' not found. Build it first: docker build -t $IMAGE ."
    fi

    # Verify cowork_net network exists (created by docker compose up -d postgres)
    if ! docker network inspect cowork_net >/dev/null 2>&1; then
        die "Network 'cowork_net' not found. Run: docker compose up -d postgres"
    fi

    # Verify postgres container is running and healthy
    local pg_status
    pg_status="$(docker inspect --format '{{.State.Health.Status}}' cowork_pg 2>/dev/null || echo "missing")"
    if [[ "$pg_status" != "healthy" ]]; then
        die "cowork_pg is not healthy (status: $pg_status). Run: docker compose up -d postgres"
    fi
}

# ---------------------------------------------------------------------------
# Sequential lock: only one task runs at a time (shared postgres constraint)
#
# Uses flock(1) on Linux; falls back to a mkdir-based atomic lock on macOS
# (where flock is not available by default).
# ---------------------------------------------------------------------------
LOCK_DIR="$DUMPS_DIR/.run.lock.d"

acquire_lock() {
    if command -v flock >/dev/null 2>&1; then
        # Linux: flock on a file descriptor
        exec 9>"$LOCK_FILE"
        if ! flock --nonblock 9 2>/dev/null; then
            warn "Another task is already running (lock: $LOCK_FILE)."
            warn "Waiting for it to finish before starting $TASK ..."
            flock 9
        fi
    else
        # macOS fallback: mkdir is atomic on POSIX filesystems
        while ! mkdir "$LOCK_DIR" 2>/dev/null; do
            warn "Another task is already running (lock: $LOCK_DIR)."
            warn "Waiting 3s before retrying ..."
            sleep 3
        done
        # Release the mkdir lock on exit (in addition to the container cleanup)
        trap 'rmdir "$LOCK_DIR" 2>/dev/null || true; cleanup' EXIT
    fi
    log "Lock acquired."
}

# ---------------------------------------------------------------------------
# Build the sanitized task tree for the agent phase (no groundtruth / evaluation)
# ---------------------------------------------------------------------------
build_staging() {
    log "Building sanitized task tree (no groundtruth/evaluation) ..."
    mkdir -p "$STAGING_DIR/tasks/finalpool/$TASK"
    # rsync the current task's source, excluding the answer key and grader code.
    rsync -a \
        --exclude 'groundtruth_workspace' \
        --exclude 'groundtruth_workspace_cn' \
        --exclude 'evaluation' \
        "$TASK_SOURCE/" "$STAGING_DIR/tasks/finalpool/$TASK/"
    # Hard guarantee: nothing answer-related leaked into the staged tree.
    if find "$STAGING_DIR/tasks" -name 'groundtruth_workspace*' -o -name 'evaluation' | grep -q .; then
        die "Sanitized staging still contains groundtruth/evaluation — aborting to avoid leak."
    fi
    log "Sanitized tree ready: $STAGING_DIR"
}

# ---------------------------------------------------------------------------
# Start an ephemeral container.
#   $1 = container name
#   $2 = host path mounted read-only at /workspace/tasks
# ---------------------------------------------------------------------------
start_container() {
    local cname="$1" tasks_src="$2"
    log "Starting container $cname ..."

    # Collect model-related env vars set on the host and forward them into the
    # container. MODEL_* configure the default runner (main.py); LLM_* are a
    # common convention a pluggable engine (AGENT_ENTRY) may read for its own
    # model client. Unset vars are skipped.
    local env_args=()
    for var in MODEL_PROVIDER MODEL_NAME MODEL_API_KEY MODEL_API_URL \
               LLM_BASE_URL LLM_API_KEY LLM_MODEL MCP_HTTP_URLS MCP_HTTP_TIMEOUT; do
        [[ -n "${!var:-}" ]] && env_args+=("-e" "${var}=${!var}")
    done

    # Always mounted: dumps (agent output), tasks (sanitized for agent / full for
    # eval — the isolation mechanism), configs (MCP server defs).
    local mount_args=(
        -v "$OUTPUT_DIR:/workspace/dumps"
        -v "$tasks_src:/workspace/tasks:ro"
        -v "$(pwd)/configs:/workspace/configs:ro"
    )
    # Harness/engine code ships INSIDE the image (Dockerfile COPY . .). By default
    # the container runs that baked code, so an end user who built the image from
    # this repo gets the fix with no host mounts. DEV_MOUNTS=1 overlays host code
    # for fast iteration: the harness (utils, run_eval.py, main.py) plus, when a
    # custom engine is plugged in, its entrypoint file.
    if [[ "${DEV_MOUNTS:-0}" == "1" ]]; then
        mount_args+=(
            -v "$(pwd)/utils:/workspace/utils:ro"
            -v "$(pwd)/main.py:/workspace/main.py:ro"
            -v "$(pwd)/scripts/run_eval.py:/workspace/scripts/run_eval.py:ro"
        )
        if [[ -n "${AGENT_ENTRY:-}" && -f "$(pwd)/${AGENT_ENTRY}" ]]; then
            mount_args+=(-v "$(pwd)/${AGENT_ENTRY}:/workspace/${AGENT_ENTRY}:ro")
        fi
    fi

    docker run -d \
        --name "$cname" \
        --network cowork_net \
        -e PGHOST=cowork_pg \
        -e PG_HOST=cowork_pg \
        -e PGPORT=5432 \
        -e PGUSER=eigent \
        -e PGPASSWORD=camel \
        -e PGDATABASE=cowork_gym \
        -e LOCAL_SERVERS_PATH=/opt/local_servers \
        -e PYTHON_BIN=/opt/venv/bin/python3 \
        ${env_args[@]+"${env_args[@]}"} \
        ${mount_args[@]+"${mount_args[@]}"} \
        -w /workspace \
        "$IMAGE" \
        sleep 3600 \
        >/dev/null

    log "Container $cname started."
}

# ---------------------------------------------------------------------------
# Wait until a container is responsive.  $1 = container name
# ---------------------------------------------------------------------------
wait_for_container() {
    local cname="$1"
    local max_wait=30
    local count=0
    log "Waiting for container $cname to be ready ..."
    while (( count < max_wait )); do
        if docker exec "$cname" true >/dev/null 2>&1; then
            log "Container $cname is ready."
            return 0
        fi
        (( count++ ))
        sleep 1
    done
    die "Container $cname did not become ready within ${max_wait}s"
}

# ---------------------------------------------------------------------------
# Run the task inside the container
# ---------------------------------------------------------------------------
# Agent entrypoint (pluggable). Default: the baseline CAMEL runner (main.py).
# Plug in a custom engine with AGENT_ENTRY=<path> (relative to /workspace). The
# engine MUST write traj_log.json + workspace into /workspace/dumps (the grader
# finds the log there); see scripts/run_eval.py.
ENTRY="${AGENT_ENTRY:-main.py}"

# Groundtruth isolation (separate agent/eval containers) needs the engine to
# support an agent-only phase (--phase agent), so the agent never shares a
# container with the groundtruth answer key. A phase-aware engine opts in with
# AGENT_PHASE_AWARE=1. Otherwise the legacy single pass runs (agent + eval
# together, full mount — NO isolation).
PHASE_SPLIT=0
if [[ "${AGENT_PHASE_AWARE:-}" == "1" ]]; then
    PHASE_SPLIT=1
fi

run_agent_phase() {
    log "Agent phase: $TASK (max_steps=$MAX_STEPS, container=$AGENT_CONTAINER) ..."

    # Fix sent_log foreign key (init.sql.gz lacks ON DELETE CASCADE)
    docker exec "$AGENT_CONTAINER" \
        /opt/venv/bin/python3 -c "
import psycopg2, os
conn = psycopg2.connect(host=os.environ['PGHOST'], database=os.environ['PGDATABASE'],
                        user=os.environ['PGUSER'], password=os.environ['PGPASSWORD'])
conn.autocommit = True
cur = conn.cursor()
try:
    cur.execute('ALTER TABLE email.sent_log DROP CONSTRAINT sent_log_message_id_fkey')
    cur.execute('ALTER TABLE email.sent_log ADD CONSTRAINT sent_log_message_id_fkey FOREIGN KEY (message_id) REFERENCES email.messages(id) ON DELETE CASCADE')
except: pass
conn.close()
" 2>/dev/null || true

    local phase_arg=()
    [[ "$PHASE_SPLIT" == "1" ]] && phase_arg=(--phase agent)
    docker exec "$AGENT_CONTAINER" \
        /opt/venv/bin/python3 "$ENTRY" \
            --task_dir  "$TASK" \
            --max_steps "$MAX_STEPS" \
            ${phase_arg[@]+"${phase_arg[@]}"} \
            --debug \
        2>&1 | tee "$OUTPUT_DIR/run.log"
}

run_eval_phase() {
    log "Eval phase: $TASK (container=$EVAL_CONTAINER, groundtruth present, standalone grader) ..."
    # Engine-independent grader: imports only the evaluator, finds the agent's
    # traj_log.json under /workspace/dumps, grades against groundtruth.
    docker exec -e PYTHONPATH=/workspace "$EVAL_CONTAINER" \
        /opt/venv/bin/python3 scripts/run_eval.py \
            --dumps_dir /workspace/dumps \
        2>&1 | tee "$OUTPUT_DIR/eval.log"
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
log "=============================================="
log "  Task:      $TASK"
log "  Max steps: $MAX_STEPS"
log "  Image:     $IMAGE"
log "  Model:     ${MODEL_NAME:-<from eval_config>} (${MODEL_PROVIDER:-<from eval_config>})"
log "  Output:    $OUTPUT_DIR"
log "=============================================="

check_prerequisites
acquire_lock

if [[ "$PHASE_SPLIT" == "1" ]]; then
    # Phase-aware engine: agent in a sanitized container (no groundtruth), then
    # the grader runs separately in a container that has the answer key.
    build_staging
    start_container "$AGENT_CONTAINER" "$STAGING_DIR/tasks"
    wait_for_container "$AGENT_CONTAINER"
    run_agent_phase
    docker stop "$AGENT_CONTAINER" >/dev/null 2>&1 || true

    start_container "$EVAL_CONTAINER" "$(pwd)/tasks"
    wait_for_container "$EVAL_CONTAINER"
    run_eval_phase
else
    # CAMEL baseline: legacy single pass (agent + eval together, full mount).
    start_container "$AGENT_CONTAINER" "$(pwd)/tasks"
    wait_for_container "$AGENT_CONTAINER"
    run_agent_phase
fi

log "Done. Results written to: $OUTPUT_DIR"
