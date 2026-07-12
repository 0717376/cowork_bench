#!/bin/bash
# Fully parallel benchmark: every task gets its own PostgreSQL + agent container.
# Concurrency is controlled by a semaphore (max N tasks running at once).
#
# Usage:
#   ./run_parallel.sh <max_concurrent> [task1 task2 ...]
#   ./run_parallel.sh 10                          # all tasks, 10 at a time
#   ./run_parallel.sh 5 task-a task-b task-c      # specific tasks
#
# Default runner: the baseline CAMEL agent (main.py). Plug in a custom engine
# with AGENT_ENTRY=<path under /workspace>; set AGENT_PHASE_AWARE=1 if it
# supports `--phase agent` / `--phase eval` to enable groundtruth isolation
# (agent runs on a sanitized task tree, grader runs separately with the answer
# key present).
#
# Environment variables:
#   MODEL_NAME / MODEL_PROVIDER / MODEL_API_KEY / MODEL_API_URL   default runner
#   LLM_BASE_URL / LLM_API_KEY / LLM_MODEL                        common custom-engine convention
#   MAX_STEPS / IMAGE / AGENT_ENTRY / AGENT_PHASE_AWARE

set -uo pipefail
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

# ─── Arguments ────────────────────────────────────────────────────────────────
MAX_CONCURRENT="${1:?Usage: $0 <max_concurrent> [task1] [task2] ...}"
shift

if [ $# -gt 0 ]; then
    TASKS=("$@")
else
    TASKS=()
    while IFS= read -r t; do TASKS+=("$t"); done < <(ls tasks/finalpool/)
fi

# ─── Config ───────────────────────────────────────────────────────────────────
# Model config comes from the standard env vars (MODEL_NAME / MODEL_PROVIDER),
# the same ones main.py reads. MODEL / PROVIDER are kept as internal aliases.
MODEL="${MODEL:-${MODEL_NAME:-}}"
PROVIDER="${PROVIDER:-${MODEL_PROVIDER:-}}"
MAX_STEPS="${MAX_STEPS:-100}"
TASK_TIMEOUT="${TASK_TIMEOUT:-1800}"   # wall-clock cap per agent phase (s); kills runner stalls
IMAGE="${IMAGE:-cowork-pack:latest}"
# Pluggable entrypoint: default baseline runner (main.py); a custom engine via
# AGENT_ENTRY. AGENT_PHASE_AWARE=1 enables groundtruth isolation (agent runs on
# a sanitized task tree, grader runs separately with the answer key present).
ENTRY="${AGENT_ENTRY:-main.py}"
PHASE_SPLIT=0; [ "${AGENT_PHASE_AWARE:-}" = "1" ] && PHASE_SPLIT=1
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
LOG_DIR="benchmark_logs/fully_parallel_${TIMESTAMP}"
DOCKER=$(which docker 2>/dev/null || echo "/usr/local/bin/docker")

mkdir -p "$LOG_DIR"

# Run manifest — exact params/knobs this run used. MODEL_QUANT / RUN_NOTES are
# free-form (e.g. quantization "fp8") since they aren't visible from the client.
python3 - "$LOG_DIR/run_meta.json" << PYEOF
import json, sys
meta = {
    "timestamp": "$TIMESTAMP",
    "log_dir": "$LOG_DIR",
    "provider": "${PROVIDER:-}",
    "model": "${MODEL:-}",
    "model_quant": "${MODEL_QUANT:-}",
    "notes": "${RUN_NOTES:-}",
    "max_concurrent": "$MAX_CONCURRENT",
    "total_tasks": ${#TASKS[@]},
    "max_steps": "$MAX_STEPS",
    "task_timeout_s": "$TASK_TIMEOUT",
    "image": "$IMAGE",
    "entry": "$ENTRY",
    "phase_split": "$PHASE_SPLIT",
    "llm": {
        "base_url": "${LLM_BASE_URL:-}",
        "param_profile": "${LLM_PARAM_PROFILE:-}",
        "or_provider": "${LLM_OR_PROVIDER:-}",
        "temperature": "${LLM_TEMPERATURE:-}",
        "top_p": "${LLM_TOP_P:-}",
        "reasoning_effort": "${LLM_REASONING_EFFORT:-}",
        "context_window": "${LLM_CONTEXT_WINDOW:-}",
        "stream_idle_timeout": "${LLM_STREAM_IDLE_TIMEOUT:-}",
        "presence_penalty": "${LLM_PRESENCE_PENALTY:-}",
        "repetition_penalty": "${LLM_REPETITION_PENALTY:-}",
        "ssl_verify": "${LLM_SSL_VERIFY:-}",
    },
}
with open(sys.argv[1], "w") as f:
    json.dump(meta, f, indent=2, ensure_ascii=False)
PYEOF

echo "============================================="
echo "Fully Parallel Benchmark"
echo "  Max concurrent: $MAX_CONCURRENT"
echo "  Total tasks:    ${#TASKS[@]}"
echo "  Model:          ${PROVIDER:-?}/${MODEL:-?}"
echo "  Max steps:      $MAX_STEPS"
echo "  Task timeout:   ${TASK_TIMEOUT}s"
echo "  Image:          $IMAGE"
echo "  Entry:          $ENTRY (phase_split=$PHASE_SPLIT)"
echo "  Log dir:        $LOG_DIR"
echo "============================================="

# ─── Verify image exists ─────────────────────────────────────────────────────
if ! $DOCKER run --rm "$IMAGE" true >/dev/null 2>&1; then
    echo "[error] Image '$IMAGE' not found or cannot run. Build it first."
    exit 1
fi

# ─── Semaphore via a FIFO ────────────────────────────────────────────────────
FIFO="$LOG_DIR/.semaphore"
mkfifo "$FIFO"
exec 3<>"$FIFO"
rm -f "$FIFO"

# Fill the semaphore with N tokens
for ((i = 0; i < MAX_CONCURRENT; i++)); do
    echo >&3
done

# ─── Summary file ────────────────────────────────────────────────────────────
SUMMARY="$LOG_DIR/summary.csv"
echo "task,status,eval_pass,duration_s" > "$SUMMARY"
SUMMARY_LOCK="$LOG_DIR/.summary.lock"

TOTAL_TASKS=${#TASKS[@]}
export TOTAL_TASKS

append_summary() {
    # Atomic append + cumulative progress line, under the same lock.
    while ! mkdir "$SUMMARY_LOCK" 2>/dev/null; do sleep 0.1; done
    echo "$1" >> "$SUMMARY"
    # Recount from disk so progress reflects the real summary state.
    local done_n pass_n eval_fail_n agent_fail_n other_n pct
    done_n=$(($(wc -l < "$SUMMARY") - 1))
    pass_n=$(awk -F, 'NR>1 && $3=="True"' "$SUMMARY" | wc -l | tr -d ' ')
    eval_fail_n=$(awk -F, 'NR>1 && $3=="False" && $2=="success"' "$SUMMARY" | wc -l | tr -d ' ')
    agent_fail_n=$(awk -F, 'NR>1 && $2!="success" && $2!="pg_fail"' "$SUMMARY" | wc -l | tr -d ' ')
    other_n=$(awk -F, 'NR>1 && $2=="pg_fail"' "$SUMMARY" | wc -l | tr -d ' ')
    if [ "$done_n" -gt 0 ]; then
        pct=$(awk -v p="$pass_n" -v d="$done_n" 'BEGIN{printf "%.1f", 100*p/d}')
    else
        pct="0.0"
    fi
    echo "[$(date +%H:%M:%S)] [progress] ${done_n}/${TOTAL_TASKS} done | pass=${pass_n} (${pct}%) | eval_fail=${eval_fail_n} | agent_fail=${agent_fail_n} | pg_fail=${other_n}"
    rmdir "$SUMMARY_LOCK"
}

# ─── Track all containers for cleanup ────────────────────────────────────────
CONTAINER_LIST="$LOG_DIR/.containers"
touch "$CONTAINER_LIST"
CONTAINER_LIST_LOCK="$LOG_DIR/.containers.lock"

register_container() {
    while ! mkdir "$CONTAINER_LIST_LOCK" 2>/dev/null; do sleep 0.1; done
    echo "$1" >> "$CONTAINER_LIST"
    rmdir "$CONTAINER_LIST_LOCK"
}

cleanup_all() {
    echo ""
    echo "Cleaning up all containers..."
    if [ -f "$CONTAINER_LIST" ]; then
        while IFS= read -r c; do
            $DOCKER rm -fv "$c" >/dev/null 2>&1 || true
        done < "$CONTAINER_LIST"
    fi
    # Close the semaphore fd
    exec 3>&- 2>/dev/null || true
    echo "Cleanup done."
}
trap cleanup_all EXIT

# ─── Export helpers for subshells ─────────────────────────────────────────────
export MODEL PROVIDER MAX_STEPS TASK_TIMEOUT IMAGE ENTRY PHASE_SPLIT DOCKER LOG_DIR SUMMARY SUMMARY_LOCK CONTAINER_LIST CONTAINER_LIST_LOCK
# Model config forwarded into the containers. MODEL_* drive the default runner;
# LLM_* are a common convention a pluggable engine (AGENT_ENTRY) may read.
export MODEL_PROVIDER="${MODEL_PROVIDER:-}"
export MODEL_NAME="${MODEL_NAME:-$MODEL}"
export MODEL_API_KEY="${MODEL_API_KEY:-}"
export MODEL_API_URL="${MODEL_API_URL:-}"
export LLM_BASE_URL="${LLM_BASE_URL:-}"
export LLM_API_KEY="${LLM_API_KEY:-}"
export LLM_MODEL="${LLM_MODEL:-}"
export LLM_CONTEXT_WINDOW="${LLM_CONTEXT_WINDOW:-}"
export LLM_STREAM_IDLE_TIMEOUT="${LLM_STREAM_IDLE_TIMEOUT:-}"
export LLM_PARAM_PROFILE="${LLM_PARAM_PROFILE:-}"
export LLM_TEMPERATURE="${LLM_TEMPERATURE:-}"
export LLM_TOP_P="${LLM_TOP_P:-}"
export LLM_MAX_TOKENS="${LLM_MAX_TOKENS:-}"
export LLM_OR_PROVIDER="${LLM_OR_PROVIDER:-}"
export LLM_OR_ALLOW_FALLBACKS="${LLM_OR_ALLOW_FALLBACKS:-}"
export LLM_SSL_VERIFY="${LLM_SSL_VERIFY:-}"
export MCP_HTTP_URLS="${MCP_HTTP_URLS:-}"
export MCP_HTTP_TIMEOUT="${MCP_HTTP_TIMEOUT:-}"
export -f append_summary register_container

# ─── Run a single task with full isolation ────────────────────────────────────
run_one_task() {
    local TASK="$1"
    local TASK_HASH=$(echo "$TASK" | md5 -q 2>/dev/null || echo "$TASK" | md5sum 2>/dev/null | cut -c1-8 || echo "$RANDOM")
    local TASK_ID="$$-${TASK_HASH:0:8}"
    local PG_CONTAINER="pg-${TASK_ID}"
    local AGENT_CONTAINER="agent-${TASK_ID}"
    local EVAL_CONTAINER="eval-${TASK_ID}"
    local TASK_LOG="$LOG_DIR/${TASK}.log"
    local NET_NAME="net-${TASK_ID}"

    local START_TS=$(date +%s)
    echo "[$(date +%H:%M:%S)] START  $TASK"

    # Create an isolated Docker network for this task
    $DOCKER network create "$NET_NAME" >> "$TASK_LOG" 2>&1 || true
    register_container "$PG_CONTAINER"
    register_container "$AGENT_CONTAINER"
    register_container "$EVAL_CONTAINER"

    # --- Start PostgreSQL ---
    $DOCKER run -d \
        --name "$PG_CONTAINER" \
        --network "$NET_NAME" \
        -e POSTGRES_DB=cowork_gym \
        -e POSTGRES_USER=eigent \
        -e POSTGRES_PASSWORD=camel \
        -v "$(pwd)/db/init.sql.gz:/docker-entrypoint-initdb.d/init.sql.gz:ro" \
        --health-cmd="pg_isready -U eigent -d cowork_gym" \
        --health-interval=3s --health-retries=20 \
        postgres:15 >> "$TASK_LOG" 2>&1

    # Wait for real TCP readiness (the docker healthcheck passes early on the
    # initdb bootstrap socket; 127.0.0.1:5432 opens only after the seed loads).
    local RETRIES=150 READY=false
    while [ $RETRIES -gt 0 ]; do
        if $DOCKER exec "$PG_CONTAINER" pg_isready -h 127.0.0.1 -p 5432 -U eigent -d cowork_gym >/dev/null 2>&1; then
            READY=true; break
        fi
        sleep 2
        RETRIES=$((RETRIES - 1))
    done

    if [ "$READY" != "true" ]; then
        echo "[$(date +%H:%M:%S)] FAIL   $TASK (postgres not healthy)" | tee -a "$TASK_LOG"
        local END_TS=$(date +%s)
        append_summary "${TASK},pg_fail,null,$((END_TS - START_TS))"
        $DOCKER rm -fv "$PG_CONTAINER" >> "$TASK_LOG" 2>&1 || true
        $DOCKER network rm "$NET_NAME" >> "$TASK_LOG" 2>&1 || true
        return 1
    fi

    # Fix sent_log foreign key
    $DOCKER run --rm --network "$NET_NAME" \
        -e PGHOST="$PG_CONTAINER" -e PGPORT=5432 \
        -e PGDATABASE=cowork_gym -e PGUSER=eigent -e PGPASSWORD=camel \
        "$IMAGE" /opt/venv/bin/python3 -c "
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
" >> "$TASK_LOG" 2>&1 || true

    # --- Forward model env into the containers ---
    # MODEL_* drive the default runner; LLM_* are a common convention a custom
    # engine (AGENT_ENTRY) may read. MODEL_PROVIDER / MODEL_NAME are set explicitly below.
    local ENV_ARGS=()
    for var in MODEL_API_KEY MODEL_API_URL \
               LLM_BASE_URL LLM_API_KEY LLM_MODEL LLM_CONTEXT_WINDOW LLM_STREAM_IDLE_TIMEOUT LLM_SSL_VERIFY \
               LLM_PARAM_PROFILE LLM_TEMPERATURE LLM_TOP_P LLM_MAX_TOKENS LLM_OR_PROVIDER LLM_OR_ALLOW_FALLBACKS \
               LLM_PRESENCE_PENALTY LLM_REPETITION_PENALTY LLM_REASONING_EFFORT \
               MCP_HTTP_URLS MCP_HTTP_TIMEOUT; do
        [ -n "${!var:-}" ] && ENV_ARGS+=("-e" "${var}=${!var}")
    done

    # --- Groundtruth isolation: sanitized task tree for the agent phase ---
    # When phase-aware, the agent runs against a copy of ONLY this task with the
    # answer key (groundtruth_workspace/ + evaluation/) stripped, shadowing the
    # full tasks dir so the agent cannot read the groundtruth. Eval runs later in
    # a separate container that mounts the full tree.
    # Code runs from the baked image; only data dirs are mounted (so each
    # container uses its own baked /workspace/.venv — no host race). The agent
    # gets a sanitized tasks tree when phase-aware (no answer key), else the full tree.
    local STAGING="" ; local AGENT_TASK_MOUNT=()
    if [ "$PHASE_SPLIT" = "1" ]; then
        STAGING=$(mktemp -d "${TMPDIR:-/tmp}/cowork-stage-${TASK_ID}.XXXXXX")
        mkdir -p "$STAGING/tasks/finalpool/$TASK"
        rsync -a --exclude groundtruth_workspace --exclude groundtruth_workspace_cn \
              --exclude evaluation \
              "$(pwd)/tasks/finalpool/$TASK/" "$STAGING/tasks/finalpool/$TASK/" >> "$TASK_LOG" 2>&1
        # Writable: preprocess mock-servers write into the throwaway staging copy.
        AGENT_TASK_MOUNT=(-v "$STAGING/tasks:/workspace/tasks")
    else
        AGENT_TASK_MOUNT=(-v "$(pwd)/tasks:/workspace/tasks:ro")
    fi

    # --- Start agent container ---
    $DOCKER run -d \
        --name "$AGENT_CONTAINER" \
        --network "$NET_NAME" \
        -e PGHOST="$PG_CONTAINER" \
        -e PG_HOST="$PG_CONTAINER" \
        -e PGPORT=5432 \
        -e PGUSER=eigent \
        -e PGPASSWORD=camel \
        -e PGDATABASE=cowork_gym \
        -e LOCAL_SERVERS_PATH=/opt/local_servers \
        -e PYTHON_BIN=/opt/venv/bin/python3 \
        -e MODEL_PROVIDER="$PROVIDER" \
        -e MODEL_NAME="$MODEL" \
        "${ENV_ARGS[@]}" \
        -v "$(pwd)/dumps:/workspace/dumps" \
        -v "$(pwd)/configs:/workspace/configs:ro" \
        "${AGENT_TASK_MOUNT[@]}" \
        -w /workspace \
        "$IMAGE" sleep 7200 >> "$TASK_LOG" 2>&1

    sleep 1

    # --- Agent phase ---
    # Model config is read from the forwarded env (MODEL_PROVIDER / MODEL_NAME or
    # LLM_*), so no --provider / --model_name args are needed.
    local PHASE_ARG=()
    [ "$PHASE_SPLIT" = "1" ] && PHASE_ARG=(--phase agent)
    $DOCKER exec \
        "$AGENT_CONTAINER" \
        timeout -k 30 "$TASK_TIMEOUT" /opt/venv/bin/python3 -u "/workspace/$ENTRY" \
            --task_dir "$TASK" \
            --max_steps "$MAX_STEPS" \
            "${PHASE_ARG[@]}" \
        >> "$TASK_LOG" 2>&1 || true

    # --- Eval phase (isolated): separate container WITH the groundtruth ---
    if [ "$PHASE_SPLIT" = "1" ]; then
        $DOCKER rm -fv "$AGENT_CONTAINER" >> "$TASK_LOG" 2>&1 || true
        $DOCKER run -d \
            --name "$EVAL_CONTAINER" \
            --network "$NET_NAME" \
            -e PGHOST="$PG_CONTAINER" \
            -e PG_HOST="$PG_CONTAINER" \
            -e PGPORT=5432 \
            -e PGUSER=eigent \
            -e PGPASSWORD=camel \
            -e PGDATABASE=cowork_gym \
            -e LOCAL_SERVERS_PATH=/opt/local_servers \
            -e PYTHON_BIN=/opt/venv/bin/python3 \
            -e MODEL_PROVIDER="$PROVIDER" \
            -e MODEL_NAME="$MODEL" \
            -v "$(pwd)/dumps:/workspace/dumps" \
            -v "$(pwd)/tasks:/workspace/tasks:ro" \
            -v "$(pwd)/configs:/workspace/configs:ro" \
            -w /workspace \
            "$IMAGE" sleep 1800 >> "$TASK_LOG" 2>&1
        sleep 1
        $DOCKER exec \
            "$EVAL_CONTAINER" \
            /opt/venv/bin/python3 -u "/workspace/$ENTRY" \
                --task_dir "$TASK" \
                --phase eval \
            >> "$TASK_LOG" 2>&1 || true
    fi

    local END_TS=$(date +%s)
    local DURATION=$((END_TS - START_TS))

    # --- Parse results ---
    local STATUS="unknown" EVAL_PASS="null"
    if grep -q "Status: success" "$TASK_LOG" 2>/dev/null; then
        STATUS="success"
    elif grep -q "Status: failed" "$TASK_LOG" 2>/dev/null; then
        STATUS="failed"
    fi

    if [ "$STATUS" = "success" ]; then
        if grep -q "Pass:.*True" "$TASK_LOG" 2>/dev/null; then
            EVAL_PASS="True"
        elif grep -q "Pass:.*False" "$TASK_LOG" 2>/dev/null; then
            EVAL_PASS="False"
        fi
    fi

    append_summary "${TASK},${STATUS},${EVAL_PASS},${DURATION}"

    # --- Determine result label ---
    local RESULT="AGENT_FAIL"
    if [ "$EVAL_PASS" = "True" ]; then
        RESULT="PASS"
    elif [ "$STATUS" = "success" ]; then
        RESULT="EVAL_FAIL"
    fi

    echo "[$(date +%H:%M:%S)] DONE   $TASK -> $RESULT (${DURATION}s)"

    # --- Cleanup this task's containers, network and staging dir ---
    $DOCKER rm -fv "$AGENT_CONTAINER" >> "$TASK_LOG" 2>&1 || true
    $DOCKER rm -fv "$EVAL_CONTAINER" >> "$TASK_LOG" 2>&1 || true
    $DOCKER rm -fv "$PG_CONTAINER" >> "$TASK_LOG" 2>&1 || true
    $DOCKER network rm "$NET_NAME" >> "$TASK_LOG" 2>&1 || true
    [ -n "$STAGING" ] && [ -d "$STAGING" ] && rm -rf "$STAGING"
}

export -f run_one_task

# ─── Launch all tasks with semaphore-controlled concurrency ───────────────────
PIDS=()
for TASK in "${TASKS[@]}"; do
    # Acquire semaphore token (blocks if all N slots are in use)
    read -u 3

    (
        run_one_task "$TASK"
        # Release semaphore token
        echo >&3
    ) &
    PIDS+=($!)
done

# ─── Wait for all tasks ──────────────────────────────────────────────────────
echo ""
echo "All ${#TASKS[@]} tasks launched (max $MAX_CONCURRENT concurrent). Waiting..."
echo ""

FAILED=0
for pid in "${PIDS[@]}"; do
    wait "$pid" || FAILED=$((FAILED + 1))
done

# ─── Report ──────────────────────────────────────────────────────────────────
echo ""
echo "============================================="
echo "RESULTS"
echo "============================================="

python3 - "$SUMMARY" "$LOG_DIR" << 'PYEOF'
import sys, csv, json, os

summary_file = sys.argv[1]
log_dir = sys.argv[2]
pass_count = eval_fail = agent_fail = other = 0
total_duration = 0
results = []

with open(summary_file) as f:
    reader = csv.DictReader(f)
    for row in reader:
        task = row["task"]
        status = row["status"]
        eval_pass = row["eval_pass"]
        duration = int(row["duration_s"])
        total_duration += duration

        if eval_pass == "True":
            label = "PASS"
            pass_count += 1
        elif status == "success":
            label = "EVAL_FAIL"
            eval_fail += 1
        elif status == "pg_fail":
            label = "PG_FAIL"
            other += 1
        else:
            label = "AGENT_FAIL"
            agent_fail += 1
        results.append((task, label, duration))

# Print individual results
for task, label, dur in sorted(results):
    print(f"  {task:<55s} {label:<12s} ({dur}s)")

total = pass_count + eval_fail + agent_fail + other
print()
if total > 0:
    print(f"  PASS:       {pass_count:4d}  ({100*pass_count/total:.1f}%)")
    print(f"  EVAL_FAIL:  {eval_fail:4d}  ({100*eval_fail/total:.1f}%)")
    print(f"  AGENT_FAIL: {agent_fail:4d}  ({100*agent_fail/total:.1f}%)")
    if other:
        print(f"  OTHER_FAIL: {other:4d}  ({100*other/total:.1f}%)")
    print(f"  TOTAL:      {total:4d}")
    print(f"  Wall time sum: {total_duration}s")
else:
    print("  No results.")

# Append aggregate to central index (one row per run; params from run_meta.json)
try:
    with open(os.path.join(log_dir, "run_meta.json")) as f:
        m = json.load(f)
except (OSError, ValueError):
    m = {}
llm = m.get("llm", {})
pass_pct = round(100 * pass_count / total, 1) if total else 0.0
index = os.path.join(os.path.dirname(log_dir) or ".", "index.csv")
cols = ["timestamp", "model", "model_quant", "reasoning_effort", "max_concurrent",
        "total", "pass", "pass_pct", "eval_fail", "agent_fail", "other",
        "idle_timeout", "context_window", "image", "notes", "log_dir"]
row = {
    "timestamp": m.get("timestamp", ""), "model": m.get("model", ""),
    "model_quant": m.get("model_quant", ""), "reasoning_effort": llm.get("reasoning_effort", ""),
    "max_concurrent": m.get("max_concurrent", ""), "total": total, "pass": pass_count,
    "pass_pct": pass_pct, "eval_fail": eval_fail, "agent_fail": agent_fail, "other": other,
    "idle_timeout": llm.get("stream_idle_timeout", ""), "context_window": llm.get("context_window", ""),
    "image": m.get("image", ""), "notes": m.get("notes", ""), "log_dir": log_dir,
}
new = not os.path.exists(index)
with open(index, "a", newline="") as f:
    w = csv.DictWriter(f, fieldnames=cols)
    if new:
        w.writeheader()
    w.writerow(row)
print(f"\n  Index: {index}")
PYEOF

echo ""
echo "Summary CSV: $SUMMARY"
echo "Task logs:   $LOG_DIR/<task>.log"
echo "Done."
