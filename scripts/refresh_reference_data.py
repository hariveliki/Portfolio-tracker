#!/usr/bin/env python3
"""Fetch reference data with yfinance primary and IBKR fallback."""

from __future__ import annotations

import argparse
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml
import yfinance as yf
from ib_insync import IB, Stock
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


@dataclass
class RefreshConfig:
    workbook: Path
    ib_host: str
    ib_port: int
    ib_client_id: int


def load_yaml_config(config_path: Path) -> dict[str, Any]:
    if not config_path.exists():
        return {}
    with config_path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def parse_args() -> RefreshConfig:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook")
    parser.add_argument("--config", default=DEFAULT_CONFIG)
    parser.add_argument("--ib-host")
    parser.add_argument("--ib-port", type=int)
    parser.add_argument("--ib-client-id", type=int)
    args = parser.parse_args()

    raw_cfg = load_yaml_config(Path(args.config))
    ib_cfg = raw_cfg.get("ib", {})
    path_cfg = raw_cfg.get("paths", {})
    return RefreshConfig(
        workbook=Path(args.workbook or path_cfg.get("workbook_output", "output/portfolio_workbook.xlsx")),
        ib_host=args.ib_host or os.getenv("IB_HOST", str(ib_cfg.get("host", "127.0.0.1"))),
        ib_port=args.ib_port or int(os.getenv("IB_PORT", str(ib_cfg.get("port", 7497)))),
        ib_client_id=args.ib_client_id or int(os.getenv("IB_CLIENT_ID", str(ib_cfg.get("client_id", 16)))),
    )


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


def _blank_row(ticker: str) -> dict[str, Any]:
    return {"Ticker": ticker, "Name": "", "Sector": "", "Industry": "", "Region": "", "Currency": "", "Beta": ""}


def _needs_ib_fallback(row: dict[str, Any]) -> bool:
    # If both identity and currency are missing, yfinance was not useful.
    return not row.get("Name") and not row.get("Currency")


def fetch_reference_yf(ticker: str) -> dict[str, Any]:
    try:
        info = yf.Ticker(ticker).info
    except Exception as exc:
        print(f"  Warning: yfinance lookup failed for {ticker}: {exc}")
        return _blank_row(ticker)

    if not isinstance(info, dict) or not info:
        return _blank_row(ticker)

    row = {
        "Ticker": ticker,
        "Name": info.get("longName") or info.get("shortName", ""),
        "Sector": info.get("sector", ""),
        "Industry": info.get("industry", ""),
        "Region": info.get("country", ""),
        "Currency": info.get("currency", ""),
        "Beta": info.get("beta") if info.get("beta") is not None else "",
    }
    return row


def fetch_reference_ibkr(ib: IB, ticker: str) -> dict[str, Any]:
    contract = Stock(ticker, "SMART", "USD")
    ib.qualifyContracts(contract)
    details = ib.reqContractDetails(contract)
    if not details:
        return _blank_row(ticker)

    detail = details[0]
    # IBKR uses CUSIP-like market names; keep mapping conservative and stable.
    region = detail.contract.primaryExchange or detail.contract.exchange or ""
    return {
        "Ticker": ticker,
        "Name": detail.longName or "",
        "Sector": detail.industry or "",
        "Industry": detail.category or detail.subcategory or "",
        "Region": region,
        "Currency": detail.contract.currency or "",
        "Beta": "",
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
    cfg = parse_args()
    workbook_path = Path(cfg.workbook)
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
    ib = IB()
    ib_connected = False
    for ticker in tickers:
        print(f"  {ticker}...")
        row = fetch_reference_yf(ticker)
        if _needs_ib_fallback(row):
            print(f"    yfinance incomplete for {ticker}; trying IBKR fallback")
            if not ib_connected:
                try:
                    ib.connect(cfg.ib_host, cfg.ib_port, clientId=cfg.ib_client_id)
                    ib_connected = True
                except Exception as exc:
                    print(f"    Warning: failed to connect to IBKR for fallback: {exc}")
            if ib_connected:
                try:
                    ib_row = fetch_reference_ibkr(ib, ticker)
                    row = {
                        "Ticker": ticker,
                        "Name": row.get("Name") or ib_row.get("Name", ""),
                        "Sector": row.get("Sector") or ib_row.get("Sector", ""),
                        "Industry": row.get("Industry") or ib_row.get("Industry", ""),
                        "Region": row.get("Region") or ib_row.get("Region", ""),
                        "Currency": row.get("Currency") or ib_row.get("Currency", ""),
                        "Beta": row.get("Beta") if row.get("Beta") not in (None, "") else ib_row.get("Beta", ""),
                    }
                except Exception as exc:
                    print(f"    Warning: IBKR fallback failed for {ticker}: {exc}")
        rows.append(row)

    if ib_connected:
        ib.disconnect()

    write_reference_table(wb, rows)
    wb.save(workbook_path)
    print(f"Updated tbl_ReferenceData for {len(tickers)} tickers in {workbook_path}")


if __name__ == "__main__":
    main()
