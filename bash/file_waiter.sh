#!/bin/bash
set -euo pipefail

readonly LOG_DIR="logs"
readonly LOG_FILE="$LOG_DIR/file_waiter.log"

mkdir -p "$LOG_DIR"

log() {
    local level=$1
    shift
    local message="$*"
    local timestamp=$(date '+%Y-%m-%d %H:%M:%S')
    
    echo "[$timestamp] [$level] $message" | tee -a "$LOG_FILE"
}

if [ $# -ne 2 ]; then
    log "ERROR" "Invalid arguments"
    echo "Usage: $0 <file_path> <timeout_seconds>" >&2
    exit 2
fi

FILE_PATH=$1
TIMEOUT=$2
START_TIME=$(date +%s)
CHECK_INTERVAL=5

while true; do
    # Текущее время
    CURRENT_TIME=$(date +%s)
    ELAPSED=$((CURRENT_TIME - START_TIME))
    
    # Проверка timeout
    if [ $ELAPSED -ge $TIMEOUT ]; then
        log "ERROR" "Timeout: file not found after ${TIMEOUT}s"
        exit 1
    fi
    
    # Проверка файла
    if [ -f "$FILE_PATH" ]; then
        log "SUCCESS" "File found: $FILE_PATH (after ${ELAPSED}s)"
        exit 0
    fi
    
    log "INFO" "Waiting... (${ELAPSED}s/${TIMEOUT}s)"
    sleep $CHECK_INTERVAL
done