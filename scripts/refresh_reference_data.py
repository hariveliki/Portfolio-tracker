#!/usr/bin/env python3
"""Fetch reference data (name, sector, industry, region, currency, beta) for each ticker via yfinance."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import yaml
import yfinance as yf
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo


DEFAULT_CONFIG = "config/portfolio.yaml"
REF_HEADERS = ["Ticker", "Name", "Sector", "Industry", "Region", "Currency", "Beta"]

HEADER_FILL = PatternFill(start_color="1C2541", end_color="1C2541", fill_type="solid")
HEADER_FONT = Font(name="Segoe UI", color="FFFFFF", bold=True, size=10)
TABLE_STYLE = TableStyleInfo(
    name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
    showRowStripes=True, showColumnStripes=False,
)


def load_yaml_config(config_path: Path) -> dict[str, Any]:
    if not config_path.exists():
        return {}
    with config_path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook")
    parser.add_argument("--config", default=DEFAULT_CONFIG)
    args = parser.parse_args()

    raw_cfg = load_yaml_config(Path(args.config))
    path_cfg = raw_cfg.get("paths", {})
    args.workbook = args.workbook or path_cfg.get("workbook_output", "output/portfolio_workbook.xlsx")
    return args


def extract_tickers(wb) -> list[str]:
    for ws in wb.worksheets:
        if "tbl_TradeLog" not in ws.tables:
            continue
        table = ws.tables["tbl_TradeLog"]
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        values = list(ws.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
            values_only=True,
        ))
        if len(values) < 2:
            return []
        headers = [str(h) for h in values[0]]
        if "Ticker" not in headers:
            return []
        idx = headers.index("Ticker")
        return sorted({
            str(r[idx]).strip().upper()
            for r in values[1:]
            if r[idx] is not None and str(r[idx]).strip() and not str(r[idx]).startswith("=")
        })
    return []


def fetch_reference(ticker: str) -> dict[str, Any]:
    try:
        info = yf.Ticker(ticker).info
    except Exception as exc:
        print(f"  Warning: yfinance lookup failed for {ticker}: {exc}")
        return {"Ticker": ticker, "Name": "", "Sector": "", "Industry": "", "Region": "", "Currency": "", "Beta": ""}

    return {
        "Ticker": ticker,
        "Name": info.get("longName") or info.get("shortName", ""),
        "Sector": info.get("sector", ""),
        "Industry": info.get("industry", ""),
        "Region": info.get("country", ""),
        "Currency": info.get("currency", ""),
        "Beta": info.get("beta") if info.get("beta") is not None else "",
    }


def write_reference_table(wb, rows: list[dict[str, Any]], table_name: str = "tbl_ReferenceData") -> None:
    target_ws = None
    for ws in wb.worksheets:
        if table_name in ws.tables:
            target_ws = ws
            break
    if target_ws is None:
        target_ws = wb["Reference_Data"] if "Reference_Data" in wb.sheetnames else wb.create_sheet("Reference_Data")

    for row in target_ws.iter_rows(
        min_row=1, max_row=max(target_ws.max_row, 1),
        min_col=1, max_col=max(target_ws.max_column, 1),
    ):
        for cell in row:
            cell.value = None
    if table_name in target_ws.tables:
        del target_ws.tables[table_name]

    for ci, h in enumerate(REF_HEADERS, 1):
        cell = target_ws.cell(row=1, column=ci, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if not rows:
        for ci in range(1, len(REF_HEADERS) + 1):
            target_ws.cell(row=2, column=ci, value="")
        last_row = 2
    else:
        for ri, row_data in enumerate(rows, 2):
            for ci, h in enumerate(REF_HEADERS, 1):
                target_ws.cell(row=ri, column=ci, value=row_data.get(h, ""))
        last_row = len(rows) + 1

    ref = f"A1:{get_column_letter(len(REF_HEADERS))}{last_row}"
    table = Table(displayName=table_name, ref=ref)
    table.tableStyleInfo = TABLE_STYLE
    target_ws.add_table(table)

    for col in target_ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        target_ws.column_dimensions[get_column_letter(col[0].column)].width = max(14, min(40, max_len + 4))
    target_ws.freeze_panes = "B2"
    target_ws.sheet_view.showGridLines = False
    target_ws.sheet_view.zoomScale = 85


def main() -> None:
    args = parse_args()
    workbook_path = Path(args.workbook)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(workbook_path)
    tickers = extract_tickers(wb)
    if not tickers:
        print("No tickers found in Trade_Log, skipping reference data refresh")
        wb.close()
        return

    print(f"Fetching reference data for {len(tickers)} tickers...")
    rows = []
    for ticker in tickers:
        print(f"  {ticker}...")
        rows.append(fetch_reference(ticker))

    write_reference_table(wb, rows)
    wb.save(workbook_path)
    print(f"Updated tbl_ReferenceData for {len(tickers)} tickers in {workbook_path}")


if __name__ == "__main__":
    main()
