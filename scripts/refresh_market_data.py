#!/usr/bin/env python3
"""Refresh market data table in the workbook from IBKR CP API (with optional yfinance fallback)."""

from __future__ import annotations

import argparse
import os
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo

from ibkr_client import IbkrClient


DEFAULT_CONFIG = "config/portfolio.yaml"


@dataclass
class RefreshConfig:
    workbook: Path
    config_path: Path
    gateway_url: str
    account_id: str
    yfinance_enabled: bool
    yfinance_default: bool
    yfinance_per_ticker: dict[str, bool]


def load_yaml_config(config_path: Path) -> dict[str, Any]:
    if not config_path.exists():
        return {}
    with config_path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def parse_args() -> RefreshConfig:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook")
    parser.add_argument("--config", default=DEFAULT_CONFIG)
    parser.add_argument("--gateway-url")
    parser.add_argument("--account-id")
    args = parser.parse_args()

    config_path = Path(args.config)
    raw_cfg = load_yaml_config(config_path)
    cpapi_cfg = raw_cfg.get("cpapi", {})
    path_cfg = raw_cfg.get("paths", {})
    fallback_cfg = raw_cfg.get("yfinance_fallback", {})

    return RefreshConfig(
        workbook=Path(args.workbook or path_cfg.get("workbook_output", "output/portfolio_workbook.xlsx")),
        config_path=config_path,
        gateway_url=args.gateway_url or os.getenv(
            "CPAPI_GATEWAY_URL",
            str(cpapi_cfg.get("gateway_url", "https://localhost:5001/v1/api")),
        ),
        account_id=args.account_id or os.getenv(
            "CPAPI_ACCOUNT_ID",
            str(cpapi_cfg.get("account_id", "")),
        ),
        yfinance_enabled=bool(fallback_cfg.get("enabled", True)),
        yfinance_default=bool(fallback_cfg.get("default", True)),
        yfinance_per_ticker={k.upper(): bool(v) for k, v in fallback_cfg.get("tickers", {}).items()},
    )


def read_table_dataframe(wb: Workbook, table_name: str) -> tuple[pd.DataFrame, Any, str]:
    for ws in wb.worksheets:
        if table_name in ws.tables:
            table = ws.tables[table_name]
            min_col, min_row, max_col, max_row = range_boundaries(table.ref)
            values = list(
                ws.iter_rows(
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                    values_only=True,
                )
            )
            if not values:
                return pd.DataFrame(), ws, table_name
            header = values[0]
            rows = [r for r in values[1:] if any(v is not None for v in r)]
            return pd.DataFrame(rows, columns=header), ws, table_name
    raise ValueError(f"Table '{table_name}' not found in workbook")


def extract_tickers_and_start(df: pd.DataFrame) -> tuple[list[str], date]:
    if "Ticker" not in df.columns:
        raise ValueError("tbl_TradeLog must have a 'Ticker' column")
    tickers = sorted({str(t).strip().upper() for t in df["Ticker"].dropna() if str(t).strip()})
    if not tickers:
        raise ValueError("No tickers found in tbl_TradeLog[Ticker]")

    start_candidates = []
    for col in ("StartDate", "TradeDate", "Date"):
        if col in df.columns:
            parsed = pd.to_datetime(df[col], errors="coerce").dropna()
            if not parsed.empty:
                start_candidates.append(parsed.min().date())
    start_date = min(start_candidates) if start_candidates else date.today().replace(year=date.today().year - 1)
    return tickers, start_date


def fetch_cpapi_series(client: IbkrClient, ticker: str, start_date: date, conid_cache: dict[str, int]) -> pd.DataFrame:
    if ticker not in conid_cache:
        conid_cache[ticker] = client.resolve_conid(ticker)
    conid = conid_cache[ticker]

    duration_days = max((date.today() - start_date).days + 5, 10)
    period = f"{min(duration_days, 1000)}d"

    data = client.get_market_history(conid, period=period, bar="1d")
    bars = data.get("data", [])
    if not bars:
        return pd.DataFrame()

    records = []
    for bar in bars:
        ts = bar.get("t")
        close = bar.get("c")
        if ts is None or close is None:
            continue
        dt = datetime.fromtimestamp(ts / 1000, tz=timezone.utc).date()
        if dt >= start_date:
            records.append({"date": dt, ticker: close})

    if not records:
        return pd.DataFrame()
    return pd.DataFrame(records)


def fetch_yf_series(ticker: str, start_date: date) -> pd.DataFrame:
    import yfinance as yf

    hist = yf.download(ticker, start=start_date.isoformat(), end=(date.today()).isoformat(), progress=False, auto_adjust=True)
    if hist.empty:
        return pd.DataFrame()
    out = hist.reset_index()[["Date", "Close"]]
    out["Date"] = pd.to_datetime(out["Date"]).dt.date
    return out.rename(columns={"Date": "date", "Close": ticker})


def should_use_yf(cfg: RefreshConfig, ticker: str) -> bool:
    if not cfg.yfinance_enabled:
        return False
    return cfg.yfinance_per_ticker.get(ticker.upper(), cfg.yfinance_default)


def write_market_table(wb: Workbook, market_df: pd.DataFrame, table_name: str = "tbl_MarketData") -> None:
    target_ws = None
    target_table = None
    for ws in wb.worksheets:
        if table_name in ws.tables:
            target_ws = ws
            target_table = ws.tables[table_name]
            break

    if target_ws is None:
        target_ws = wb["MarketData"] if "MarketData" in wb.sheetnames else wb.create_sheet("MarketData")

    for row in target_ws.iter_rows(min_row=1, max_row=target_ws.max_row, min_col=1, max_col=max(target_ws.max_column, 1)):
        for cell in row:
            cell.value = None

    if market_df.empty:
        headers = ["Date"]
        rows = []
    else:
        headers = ["Date", *[c for c in market_df.columns if c != "date"]]
        rows = []
        for _, r in market_df.iterrows():
            rows.append([r["date"], *[r[c] for c in headers[1:]]])

    for i, h in enumerate(headers, start=1):
        target_ws.cell(row=1, column=i, value=h)
    for ridx, row in enumerate(rows, start=2):
        for cidx, v in enumerate(row, start=1):
            target_ws.cell(row=ridx, column=cidx, value=v)

    max_row = max(len(rows) + 1, 2)
    max_col = len(headers)
    ref = f"A1:{get_column_letter(max_col)}{max_row}"

    if target_table is None:
        target_table = Table(displayName=table_name, ref=ref)
        target_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
        target_ws.add_table(target_table)
    else:
        target_table.ref = ref


def main() -> None:
    cfg = parse_args()
    if not cfg.workbook.exists():
        raise FileNotFoundError(f"Workbook not found: {cfg.workbook}")

    wb = load_workbook(cfg.workbook)
    trade_df, _, _ = read_table_dataframe(wb, "tbl_TradeLog")
    tickers, start_date = extract_tickers_and_start(trade_df)

    conid_cache: dict[str, int] = {}
    merged: pd.DataFrame | None = None

    with IbkrClient(base_url=cfg.gateway_url, account_id=cfg.account_id) as client:
        client.check_auth()

        for ticker in tickers:
            series = pd.DataFrame()
            try:
                series = fetch_cpapi_series(client, ticker, start_date, conid_cache)
            except Exception as exc:
                print(f"CP API fetch failed for {ticker}: {exc}")

            if series.empty and should_use_yf(cfg, ticker):
                print(f"Using yfinance fallback for {ticker}")
                series = fetch_yf_series(ticker, start_date)

            if series.empty:
                print(f"No market data fetched for {ticker}")
                continue

            merged = series if merged is None else merged.merge(series, on="date", how="outer")

    if merged is None:
        merged = pd.DataFrame(columns=["date"])
    merged = merged.sort_values("date").reset_index(drop=True)

    write_market_table(wb, merged)
    wb.save(cfg.workbook)
    print(f"Updated tbl_MarketData for {len(tickers)} tickers in {cfg.workbook}")


if __name__ == "__main__":
    main()
