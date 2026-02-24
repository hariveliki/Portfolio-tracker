#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_DIR="${ROOT_DIR}/logs"
mkdir -p "${LOG_DIR}"

if [[ ! -d "${ROOT_DIR}/.venv" ]]; then
  echo "Missing virtual environment at ${ROOT_DIR}/.venv" >&2
  exit 1
fi

# shellcheck disable=SC1091
source "${ROOT_DIR}/.venv/bin/activate"

run_step() {
  local name="$1"
  shift
  local ts
  ts="$(date +"%Y%m%d_%H%M%S")"
  local logfile="${LOG_DIR}/${ts}_${name}.log"

  echo "Running ${name}..."
  if "$@" >"${logfile}" 2>&1; then
    echo "${name} completed. Log: ${logfile}"
  else
    local code=$?
    echo "${name} failed (exit ${code}). See log: ${logfile}" >&2
    tail -n 50 "${logfile}" >&2 || true
    exit "${code}"
  fi
}

WORKBOOK_PATH="$("${ROOT_DIR}/.venv/bin/python" - <<'PY'
from pathlib import Path
import yaml

root = Path.cwd()
config_path = root / "config" / "portfolio.yaml"
raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) if config_path.exists() else {}
workbook = Path((raw.get("paths", {}) or {}).get("workbook_output", "output/portfolio_workbook.xlsx"))
if not workbook.is_absolute():
    workbook = (root / workbook).resolve()
print(workbook)
PY
)"

if [[ ! -f "${WORKBOOK_PATH}" ]]; then
  run_step build_workbook python "${ROOT_DIR}/scripts/build_workbook.py"
fi

run_step sync_ibkr_trades python "${ROOT_DIR}/scripts/sync_ibkr_trades.py"
run_step refresh_market_data python "${ROOT_DIR}/scripts/refresh_market_data.py"
run_step refresh_reference_data python "${ROOT_DIR}/scripts/refresh_reference_data.py"
run_step rebuild_analytics python "${ROOT_DIR}/scripts/rebuild_analytics.py"
run_step reconcile python "${ROOT_DIR}/scripts/reconcile.py"

echo "Portfolio update completed successfully."
