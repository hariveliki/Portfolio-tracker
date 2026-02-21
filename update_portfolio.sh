#!/usr/bin/env bash
set -euo pipefail

python3 scripts/ibkr_smoketest.py
python3 scripts/sync_ibkr_trades.py
python3 scripts/refresh_market_data.py
python3 scripts/build_workbook.py
python3 scripts/reconcile.py
