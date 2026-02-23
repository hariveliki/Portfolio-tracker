#!/usr/bin/env python3
"""Sync IBKR executions into the TradeLog workbook."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import yaml
from ib_insync import ExecutionFilter, IB
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table

RAW_COLUMNS = [
    "IB_ExecId",
    "Date",
    "Ticker",
    "Action",
    "Shares",
    "Price",
    "Commission",
    "permId",
    "account",
    "exchange",
    "currency",
]
FORMULA_COLUMNS = ["Gross_Value", "Net_Cost", "Signed_Shares", "Signed_Cash", "Cash_Balance"]


@dataclass
class SyncConfig:
    host: str
    port: int
    client_id: int
    workbook_path: Path


def load_config(config_path: Path) -> SyncConfig:
    with config_path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}

    ib = raw.get("ib", {})
    paths = raw.get("paths", {})

    return SyncConfig(
        host=str(ib.get("host", "127.0.0.1")),
        port=int(ib.get("port", 7497)),
        client_id=int(ib.get("client_id", 41)),
        workbook_path=Path(paths.get("workbook_output", "output/portfolio_workbook.xlsx")),
    )


def load_state(state_path: Path) -> dict[str, Any]:
    if not state_path.exists():
        return {}
    with state_path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_state(state_path: Path, state: dict[str, Any]) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    with state_path.open("w", encoding="utf-8") as handle:
        json.dump(state, handle, indent=2, sort_keys=True)
        handle.write("\n")


def format_execution_filter_time(last_sync_utc: str | None) -> str | None:
    if not last_sync_utc:
        return None
    stamp = datetime.fromisoformat(last_sync_utc.replace("Z", "+00:00")).astimezone(UTC)
    return stamp.strftime("%Y%m%d %H:%M:%S")


def normalize_trading_date(fill: Any) -> str:
    candidates = [
        getattr(fill.execution, "time", None),
        getattr(fill.execution, "execTime", None),
        getattr(fill, "time", None),
    ]
    for candidate in candidates:
        if not candidate:
            continue
        if isinstance(candidate, datetime):
            return candidate.astimezone(UTC).date().isoformat()
        try:
            parsed = datetime.fromisoformat(str(candidate).replace("Z", "+00:00"))
            if parsed.tzinfo is None:
                return parsed.date().isoformat()
            return parsed.astimezone(UTC).date().isoformat()
        except ValueError:
            continue
    return datetime.now(tz=UTC).date().isoformat()


def find_trade_table(workbook_path: Path, table_name: str = "tbl_TradeLog") -> tuple[Any, Worksheet, Table]:
    wb = load_workbook(workbook_path)
    for ws in wb.worksheets:
        for name, table in ws.tables.items():
            if name == table_name:
                return wb, ws, table
    wb.close()
    raise RuntimeError(f"Could not find table {table_name!r} in {workbook_path}")


def read_table_headers(ws: Worksheet, table: Table) -> tuple[list[str], int, int, int, int]:
    start, end = table.ref.split(":")
    start_col = ws[start].column
    start_row = ws[start].row
    end_col = ws[end].column
    end_row = ws[end].row
    headers = [ws.cell(row=start_row, column=col).value for col in range(start_col, end_col + 1)]
    return [str(h) if h is not None else "" for h in headers], start_col, start_row, end_col, end_row


def existing_exec_ids(ws: Worksheet, headers: list[str], start_col: int, start_row: int, end_row: int) -> set[str]:
    if "IB_ExecId" not in headers:
        raise RuntimeError("tbl_TradeLog must include IB_ExecId column")
    idx = headers.index("IB_ExecId")
    col = start_col + idx
    values: set[str] = set()
    for row in range(start_row + 1, end_row + 1):
        cell_value = ws.cell(row=row, column=col).value
        if cell_value is not None and str(cell_value).strip():
            values.add(str(cell_value).strip())
    return values


def map_fill(fill: Any) -> dict[str, Any]:
    commission_report = getattr(fill, "commissionReport", None)
    execution = fill.execution
    contract = fill.contract
    return {
        "IB_ExecId": execution.execId,
        "Date": normalize_trading_date(fill),
        "Ticker": contract.symbol,
        "Action": execution.side,
        "Shares": execution.shares,
        "Price": execution.price,
        "Commission": getattr(commission_report, "commission", 0) if commission_report else 0,
        "permId": getattr(execution, "permId", None),
        "account": getattr(execution, "acctNumber", None),
        "exchange": getattr(execution, "exchange", None),
        "currency": getattr(contract, "currency", None),
    }


def copy_formula_columns(
    ws: Worksheet,
    headers: list[str],
    start_col: int,
    template_row: int,
    first_new_row: int,
    last_new_row: int,
) -> None:
    for col_name in FORMULA_COLUMNS:
        if col_name not in headers:
            continue
        col = start_col + headers.index(col_name)
        template_cell = ws.cell(row=template_row, column=col)
        formula = template_cell.value
        if not isinstance(formula, str) or not formula.startswith("="):
            continue
        template_coordinate = template_cell.coordinate
        for row in range(first_new_row, last_new_row + 1):
            new_cell = ws.cell(row=row, column=col)
            new_cell.value = Translator(formula, origin=template_coordinate).translate_formula(new_cell.coordinate)


def append_rows(workbook_path: Path, rows: list[dict[str, Any]]) -> int:
    if not rows:
        return 0

    wb, ws, table = find_trade_table(workbook_path)
    headers, start_col, start_row, end_col, end_row = read_table_headers(ws, table)

    first_new_row = end_row + 1
    for row_idx, row_data in enumerate(rows, start=first_new_row):
        for col_offset, header in enumerate(headers):
            col = start_col + col_offset
            if header in RAW_COLUMNS:
                ws.cell(row=row_idx, column=col).value = row_data.get(header)

    copy_formula_columns(
        ws,
        headers,
        start_col,
        template_row=end_row,
        first_new_row=first_new_row,
        last_new_row=first_new_row + len(rows) - 1,
    )

    table.ref = f"{ws.cell(row=start_row, column=start_col).coordinate}:{ws.cell(row=first_new_row + len(rows) - 1, column=end_col).coordinate}"
    wb.save(workbook_path)
    wb.close()
    return len(rows)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    config = load_config(root / "config" / "portfolio.yaml")
    state_path = root / "state" / "state.json"
    state = load_state(state_path)

    ib = IB()
    max_seen_ts: datetime | None = None

    try:
        ib.connect(config.host, config.port, clientId=config.client_id, readonly=True)

        filter_time = format_execution_filter_time(state.get("last_sync_utc"))
        execution_filter = ExecutionFilter(time=filter_time) if filter_time else ExecutionFilter()

        fills = ib.reqExecutions(execution_filter)
        mapped = [map_fill(fill) for fill in fills]

        wb, ws, table = find_trade_table(config.workbook_path)
        headers, start_col, start_row, _, end_row = read_table_headers(ws, table)
        existing_ids = existing_exec_ids(ws, headers, start_col, start_row, end_row)
        wb.close()

        new_rows: list[dict[str, Any]] = []
        for fill, row in zip(fills, mapped):
            exec_id = str(row["IB_ExecId"])
            if exec_id in existing_ids:
                continue
            new_rows.append(row)
            exec_time = getattr(fill.execution, "time", None)
            if isinstance(exec_time, datetime):
                time_obj = exec_time if exec_time.tzinfo else exec_time.replace(tzinfo=UTC)
                max_seen_ts = time_obj.astimezone(UTC) if not max_seen_ts else max(max_seen_ts, time_obj.astimezone(UTC))

        inserted = append_rows(config.workbook_path, new_rows)

        if inserted > 0:
            state["last_sync_utc"] = (max_seen_ts or datetime.now(tz=UTC)).isoformat().replace("+00:00", "Z")
            save_state(state_path, state)
    finally:
        if ib.isConnected():
            ib.disconnect()


if __name__ == "__main__":
    main()
