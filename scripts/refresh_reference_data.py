#!/usr/bin/env python3
"""Fetch reference data with yfinance primary and IBKR fallback."""

from __future__ import annotations

import argparse
import io
import os
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import yaml
import yfinance as yf
from ib_insync import IB, Stock
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo


DEFAULT_CONFIG = "config/portfolio.yaml"
REF_HEADERS = ["Ticker", "Name", "Sector", "Industry", "Region", "Currency", "Beta"]
EXPOSURE_HEADERS = ["Ticker", "Dimension", "Bucket", "Weight"]

# Yahoo/iShares sector names -> GICS names; keys are lower-case without spaces/underscores.
YF_TO_GICS = {
    "realestate": "Real Estate", "consumercyclical": "Consumer Discretionary",
    "basicmaterials": "Materials", "consumerdefensive": "Consumer Staples",
    "technology": "Information Technology", "financialservices": "Financials",
    "healthcare": "Health Care", "communication": "Communication Services",
}

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
    instruments: dict[str, dict[str, Any]]


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
        instruments={str(k).upper(): v or {} for k, v in (raw_cfg.get("instruments") or {}).items()},
    )


def extract_tickers(wb) -> dict[str, tuple[str, str]]:
    """Return {ticker: (exchange, currency)} from the latest Trade_Log row per ticker."""
    for ws in wb.worksheets:
        if "tbl_TradeLog" not in ws.tables:
            continue
        min_col, min_row, max_col, max_row = range_boundaries(ws.tables["tbl_TradeLog"].ref)
        values = list(ws.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
            values_only=True,
        ))
        if len(values) < 2:
            return {}
        headers = [str(h) for h in values[0]]
        if "Ticker" not in headers:
            return {}
        idx = headers.index("Ticker")
        ex_idx = headers.index("exchange") if "exchange" in headers else None
        ccy_idx = headers.index("currency") if "currency" in headers else None

        def text(row, i):
            return str(row[i] or "").strip().upper() if i is not None else ""

        hints: dict[str, tuple[str, str]] = {}
        for r in values[1:]:
            ticker = str(r[idx] or "").strip().upper()
            if not ticker or ticker.startswith("="):
                continue
            hints[ticker] = (text(r, ex_idx), text(r, ccy_idx))
        return dict(sorted(hints.items()))
    return {}


def _blank_row(ticker: str) -> dict[str, Any]:
    return {"Ticker": ticker, "Name": "", "Sector": "", "Industry": "", "Region": "", "Currency": "", "Beta": ""}


def fetch_reference_yf(ticker: str, symbol: str) -> dict[str, Any]:
    try:
        info = yf.Ticker(symbol).info
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


def gics(name: str) -> str:
    key = name.lower().replace(" ", "").replace("_", "")
    return YF_TO_GICS.get(key, name)


def fetch_ishares_weights(url: str) -> dict[str, dict[str, float]]:
    """Sector and region weights (%) summed from an iShares holdings CSV."""
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=60) as response:
        lines = response.read().decode("utf-8-sig").splitlines()
    start = next(i for i, line in enumerate(lines) if line.startswith("Ticker,"))
    df = pd.read_csv(io.StringIO("\n".join(lines[start:])), usecols=["Sector", "Location", "Weight (%)"])
    df["Weight (%)"] = pd.to_numeric(df["Weight (%)"], errors="coerce")
    df = df[df["Weight (%)"] > 0]
    df["Sector"] = df["Sector"].map(gics)
    return {
        "Sector": df.groupby("Sector")["Weight (%)"].sum().to_dict(),
        "Region": df.groupby("Location")["Weight (%)"].sum().to_dict(),
    }


def fetch_fund_sectors(symbol: str) -> dict[str, float]:
    """Sector weights of an ETF as Yahoo reports them; empty for stocks and bond funds."""
    try:
        weights = yf.Ticker(symbol).funds_data.sector_weightings or {}
    except Exception:
        return {}
    return {gics(k.replace("_", " ").title()): float(v) for k, v in weights.items() if v}


def build_exposure(ticker: str, row: dict[str, Any], symbol: str, override: dict[str, Any]) -> list[dict[str, Any]]:
    """Look-through rows (Ticker, Dimension, Bucket, Weight); every ticker sums to 1 per dimension."""
    fund: dict[str, dict[str, float]] = {}
    if override.get("holdings_url"):
        try:
            fund = fetch_ishares_weights(override["holdings_url"])
        except Exception as exc:
            print(f"    Warning: iShares holdings failed for {ticker}: {exc}")
    sectors = (override.get("sectors") or fund.get("Sector")
               or ({gics(row["Sector"]): 1.0} if row["Sector"] else fetch_fund_sectors(symbol)))
    regions = override.get("regions") or fund.get("Region") or ({row["Region"]: 1.0} if row["Region"] else {})
    rows = []
    for dimension, weights in (("Sector", sectors), ("Region", regions)):
        total = sum(weights.values())
        if total <= 0:
            print(f"    Warning: no {dimension.lower()} data for {ticker}")
            weights, total = {"Unclassified": 1.0}, 1.0
        rows += [{"Ticker": ticker, "Dimension": dimension, "Bucket": b, "Weight": w / total}
                 for b, w in weights.items()]
    return rows


def fetch_reference_ibkr(ib: IB, ticker: str, exchange: str, currency: str) -> dict[str, Any]:
    """Name and currency as IBKR shows them; the contract is looked up via the Trade_Log hint."""
    candidates = [(exchange or "SMART", currency or "USD"), ("SMART", currency or "USD"), ("SMART", "USD")]
    for ex, ccy in dict.fromkeys(candidates):
        details = ib.reqContractDetails(Stock(ticker, ex, ccy))
        if details:
            return {"Name": details[0].longName or "", "Currency": details[0].contract.currency or ""}
    return {}


def write_table(wb, rows: list[dict[str, Any]], table_name: str, sheet_name: str, headers: list[str]) -> None:
    target_ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)

    for row in target_ws.iter_rows(
        min_row=1, max_row=max(target_ws.max_row, 1),
        min_col=1, max_col=max(target_ws.max_column, 1),
    ):
        for cell in row:
            cell.value = None
    if table_name in target_ws.tables:
        del target_ws.tables[table_name]

    for ci, h in enumerate(headers, 1):
        cell = target_ws.cell(row=1, column=ci, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if not rows:
        for ci in range(1, len(headers) + 1):
            target_ws.cell(row=2, column=ci, value="")
        last_row = 2
    else:
        for ri, row_data in enumerate(rows, 2):
            for ci, h in enumerate(headers, 1):
                target_ws.cell(row=ri, column=ci, value=row_data.get(h, ""))
        last_row = len(rows) + 1

    ref = f"A1:{get_column_letter(len(headers))}{last_row}"
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
    hints = extract_tickers(wb)
    if not hints:
        print("No tickers found in Trade_Log, skipping reference data refresh")
        wb.close()
        return

    print(f"Fetching reference data for {len(hints)} tickers...")
    ib = IB()
    try:
        ib.connect(cfg.ib_host, cfg.ib_port, clientId=cfg.ib_client_id)
    except Exception as exc:
        print(f"  Warning: IBKR unavailable, names will come from yfinance only: {exc}")

    rows, exposure = [], []
    for ticker, (exchange, currency) in hints.items():
        print(f"  {ticker}...")
        override = cfg.instruments.get(ticker, {})
        symbol = override.get("yahoo", ticker)
        row = fetch_reference_yf(ticker, symbol)
        if ib.isConnected():
            try:
                row.update({k: v for k, v in fetch_reference_ibkr(ib, ticker, exchange, currency).items() if v})
            except Exception as exc:
                print(f"    Warning: IBKR lookup failed for {ticker}: {exc}")
        rows.append(row)
        exposure += build_exposure(ticker, row, symbol, override)

    if ib.isConnected():
        ib.disconnect()

    write_table(wb, rows, "tbl_ReferenceData", "Reference_Data", REF_HEADERS)
    write_table(wb, exposure, "tbl_ReferenceExposure", "Reference_Exposure", EXPOSURE_HEADERS)
    wb.save(workbook_path)
    print(f"Updated tbl_ReferenceData for {len(hints)} tickers in {workbook_path}")


if __name__ == "__main__":
    main()
