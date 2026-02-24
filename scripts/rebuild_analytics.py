#!/usr/bin/env python3
"""Rebuild all analytics tables (Daily_Units, Daily_Cash, Positions, Equity_Curve,
Sector/Geographic Exposure, Dashboard) from the data-lake sheets."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
import yaml
from openpyxl import load_workbook
from openpyxl.chart import LineChart, PieChart, Reference as ChartRef
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo

DEFAULT_CONFIG = "config/portfolio.yaml"

HEADER_FILL = PatternFill(start_color="1C2541", end_color="1C2541", fill_type="solid")
HEADER_FONT = Font(name="Segoe UI", color="FFFFFF", bold=True, size=10)
GREEN_FONT = Font(name="Segoe UI", color="007A33", size=10)
RED_FONT = Font(name="Segoe UI", color="B81D13", size=10)
GREEN_FILL = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
RED_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
KPI_LABEL_FONT = Font(name="Segoe UI", bold=True, size=11, color="1C2541")
KPI_VALUE_FONT = Font(name="Segoe UI", size=11)
TITLE_FONT = Font(name="Segoe UI", bold=True, size=14, color="1C2541")

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


def _read_table_df(wb, table_name: str) -> pd.DataFrame:
    for ws in wb.worksheets:
        if table_name not in ws.tables:
            continue
        table = ws.tables[table_name]
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        values = list(ws.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
            values_only=True,
        ))
        if len(values) < 2:
            return pd.DataFrame()
        headers = [str(h) for h in values[0]]
        rows = [
            r for r in values[1:]
            if any(
                v is not None and str(v).strip() and not str(v).startswith("=")
                for v in r
            )
        ]
        return pd.DataFrame(rows, columns=headers)
    return pd.DataFrame()


def _read_setting(wb, param_name: str, default: Any = None) -> Any:
    if "Settings" not in wb.sheetnames:
        return default
    ws = wb["Settings"]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2, values_only=True):
        if row[0] and str(row[0]).strip() == param_name:
            return row[1]
    return default


def _clear_sheet(ws, max_col_override: int | None = None) -> None:
    max_col = max_col_override or max(ws.max_column, 1)
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 1), min_col=1, max_col=max_col):
        for cell in row:
            cell.value = None
    for tname in list(ws.tables.keys()):
        del ws.tables[tname]
    for merge in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merge))
    ws.conditional_formatting._cf_rules = []


def _write_table(ws, table_name: str, headers: list[str], rows: list[list],
                 number_formats: dict[int, str] | None = None) -> None:
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if not rows:
        for ci in range(1, len(headers) + 1):
            ws.cell(row=2, column=ci, value="")
        last_row = 2
    else:
        for ri, row_data in enumerate(rows, 2):
            for ci, val in enumerate(row_data, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                if number_formats and ci in number_formats:
                    cell.number_format = number_formats[ci]
        last_row = len(rows) + 1

    ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    table = Table(displayName=table_name, ref=ref)
    table.tableStyleInfo = TABLE_STYLE
    ws.add_table(table)

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(14, min(40, max_len + 4))

    ws.freeze_panes = "B2"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85


def _apply_pos_neg_cf(ws, col_letters: list[str], max_row: int) -> None:
    for letter in col_letters:
        rng = f"{letter}2:{letter}{max_row}"
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL, font=GREEN_FONT),
        )
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL, font=RED_FONT),
        )


def _prepare_trades(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame(columns=["Date", "Ticker", "Shares", "Price", "Commission", "Action",
                                     "Signed_Shares", "Signed_Cash"])
    df = raw.copy()
    df["Shares"] = pd.to_numeric(df["Shares"], errors="coerce").fillna(0)
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce").fillna(0)
    df["Commission"] = pd.to_numeric(df["Commission"], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Date"] = df["Date"].dt.date
    df["Ticker"] = df["Ticker"].astype(str).str.strip().str.upper()
    df = df[df["Ticker"].ne("")]
    df["Signed_Shares"] = df.apply(
        lambda r: r["Shares"] if str(r["Action"]).upper() == "BUY" else -r["Shares"], axis=1,
    )
    df["Signed_Cash"] = -df["Signed_Shares"] * df["Price"] - df["Commission"]
    return df


def _prepare_market(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame(columns=["Date", "Ticker", "Close"])
    df = raw.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Date"] = df["Date"].dt.date
    df["Ticker"] = df["Ticker"].astype(str).str.strip().str.upper()
    df["Close"] = pd.to_numeric(df["Close"], errors="coerce")
    return df


def _prepare_reference(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame(columns=["Ticker", "Name", "Sector", "Industry", "Region", "Currency", "Beta"])
    df = raw.copy()
    df["Ticker"] = df["Ticker"].astype(str).str.strip().str.upper()
    return df


def compute_daily_units(trades: pd.DataFrame, trading_dates: list[date]) -> pd.DataFrame:
    if trades.empty or not trading_dates:
        return pd.DataFrame(columns=["Date", "Ticker", "Units"])

    trade_events = trades.groupby(["Date", "Ticker"])["Signed_Shares"].sum().reset_index()
    tickers = sorted(trade_events["Ticker"].unique())

    results: list[dict] = []
    for ticker in tickers:
        t_events = trade_events[trade_events["Ticker"] == ticker].set_index("Date")["Signed_Shares"]
        cumulative = 0.0
        for d in trading_dates:
            if d in t_events.index:
                cumulative += t_events[d]
            if cumulative != 0:
                results.append({"Date": d, "Ticker": ticker, "Units": cumulative})

    return pd.DataFrame(results)


def compute_daily_cash(trades: pd.DataFrame, trading_dates: list[date], start_capital: float) -> pd.DataFrame:
    if not trading_dates:
        return pd.DataFrame(columns=["Date", "Cash_Balance"])

    cash_events = trades.groupby("Date")["Signed_Cash"].sum() if not trades.empty else pd.Series(dtype=float)

    balance = start_capital
    rows: list[dict] = []
    for d in trading_dates:
        if d in cash_events.index:
            balance += cash_events[d]
        rows.append({"Date": d, "Cash_Balance": balance})
    return pd.DataFrame(rows)


def compute_positions(trades: pd.DataFrame, market: pd.DataFrame,
                      reference: pd.DataFrame) -> pd.DataFrame:
    if trades.empty:
        return pd.DataFrame()

    current_units = trades.groupby("Ticker")["Signed_Shares"].sum()
    current_units = current_units[current_units.abs() > 1e-9]
    if current_units.empty:
        return pd.DataFrame()

    avg_cost: dict[str, float] = {}
    for ticker in current_units.index:
        t = trades[trades["Ticker"] == ticker]
        total_cost = (t["Signed_Shares"] * t["Price"] + t["Commission"]).sum()
        total_shares = t["Signed_Shares"].sum()
        avg_cost[ticker] = abs(total_cost / total_shares) if total_shares != 0 else 0

    ref_map = {}
    if not reference.empty:
        for _, r in reference.iterrows():
            ref_map[r["Ticker"]] = r

    latest_prices: dict[str, float] = {}
    day_changes: dict[str, float] = {}
    if not market.empty:
        for ticker in current_units.index:
            t_market = market[market["Ticker"] == ticker].sort_values("Date")
            if t_market.empty:
                continue
            latest_prices[ticker] = float(t_market["Close"].iloc[-1])
            if len(t_market) >= 2:
                prev = float(t_market["Close"].iloc[-2])
                curr = float(t_market["Close"].iloc[-1])
                day_changes[ticker] = (curr / prev - 1) if prev != 0 else 0
            else:
                day_changes[ticker] = 0

    rows = []
    for ticker in sorted(current_units.index):
        units = current_units[ticker]
        ref = ref_map.get(ticker, {})
        last_price = latest_prices.get(ticker, 0)
        rows.append({
            "Ticker": ticker,
            "Company": ref.get("Name", "") if isinstance(ref, (dict, pd.Series)) else "",
            "Sector": ref.get("Sector", "") if isinstance(ref, (dict, pd.Series)) else "",
            "Region": ref.get("Region", "") if isinstance(ref, (dict, pd.Series)) else "",
            "Units": units,
            "Avg_Cost": avg_cost.get(ticker, 0),
            "Last_Price": last_price,
            "Day_Change_Pct": day_changes.get(ticker, 0),
        })
    return pd.DataFrame(rows)


def compute_equity_curve(daily_units: pd.DataFrame, daily_cash: pd.DataFrame,
                         market: pd.DataFrame) -> pd.DataFrame:
    if daily_cash.empty:
        return pd.DataFrame(columns=["Date", "Portfolio_Value", "Cash_Balance"])

    trading_dates = sorted(daily_cash["Date"].unique())

    market_wide = market.pivot_table(index="Date", columns="Ticker", values="Close") if not market.empty else pd.DataFrame()
    units_wide = daily_units.pivot_table(index="Date", columns="Ticker", values="Units", fill_value=0) if not daily_units.empty else pd.DataFrame()

    cash_lookup = dict(zip(daily_cash["Date"], daily_cash["Cash_Balance"]))

    rows: list[dict] = []
    for d in trading_dates:
        pv = 0.0
        if d in units_wide.index and d in market_wide.index:
            u_row = units_wide.loc[d]
            p_row = market_wide.loc[d]
            common = u_row.index.intersection(p_row.index)
            pv = float((u_row[common] * p_row[common]).sum())

        rows.append({
            "Date": d,
            "Portfolio_Value": round(pv, 2),
            "Cash_Balance": round(cash_lookup.get(d, 0), 2),
        })

    return pd.DataFrame(rows)


def _write_daily_units(wb, daily_units: pd.DataFrame) -> None:
    ws = wb["Daily_Units"]
    _clear_sheet(ws)
    rows = [[r["Date"], r["Ticker"], r["Units"]] for _, r in daily_units.iterrows()] if not daily_units.empty else []
    _write_table(ws, "tbl_DailyUnits", ["Date", "Ticker", "Units"], rows)


def _write_daily_cash(wb, daily_cash: pd.DataFrame) -> None:
    ws = wb["Daily_Cash"]
    _clear_sheet(ws)
    rows = [[r["Date"], round(r["Cash_Balance"], 2)] for _, r in daily_cash.iterrows()] if not daily_cash.empty else []
    _write_table(ws, "tbl_DailyCash", ["Date", "Cash_Balance"], rows,
                 number_formats={2: "#,##0.00"})


def _write_positions(wb, positions: pd.DataFrame) -> None:
    ws = wb["Positions"]
    _clear_sheet(ws)

    headers = [
        "Ticker", "Company", "Sector", "Region", "Units", "Avg_Cost", "Last_Price",
        "Day_Change_Pct", "Cost_Basis", "Market_Value", "Unrealized_PnL", "PnL_Pct", "Weight",
    ]
    rows = []
    for _, p in positions.iterrows():
        rows.append([
            p["Ticker"], p["Company"], p["Sector"], p["Region"],
            p["Units"], round(p["Avg_Cost"], 4), round(p["Last_Price"], 4),
            p.get("Day_Change_Pct", 0),
            "=[@Units]*[@Avg_Cost]",
            "=[@Units]*[@Last_Price]",
            "=[@Market_Value]-[@Cost_Basis]",
            "=IFERROR([@Unrealized_PnL]/[@Cost_Basis],0)",
            "=IFERROR([@Market_Value]/SUM(tbl_Positions[Market_Value]),0)",
        ])

    fmt = {5: "#,##0.00", 6: "#,##0.00", 7: "#,##0.00", 8: "0.00%",
           9: "#,##0.00", 10: "#,##0.00", 11: "#,##0.00", 12: "0.00%", 13: "0.00%"}
    _write_table(ws, "tbl_Positions", headers, rows, number_formats=fmt)

    if rows:
        _apply_pos_neg_cf(ws, ["H", "K", "L"], len(rows) + 1)


def _write_equity_curve(wb, eq_curve: pd.DataFrame) -> None:
    ws = wb["Equity_Curve"]
    _clear_sheet(ws)

    headers = [
        "Date", "Portfolio_Value", "Cash_Balance", "NAV", "Daily_Return",
        "Cumulative_Return", "Peak_NAV", "Drawdown",
    ]
    rows = []
    for _, r in eq_curve.iterrows():
        rows.append([
            r["Date"],
            r["Portfolio_Value"],
            r["Cash_Balance"],
            "=[@Portfolio_Value]+[@Cash_Balance]",
            "=IFERROR([@NAV]/OFFSET([@NAV],-1,0)-1,0)",
            "=IFERROR([@NAV]/INDEX(tbl_EquityCurve[NAV],1)-1,0)",
            "=MAX(INDEX(tbl_EquityCurve[NAV],1):[@NAV])",
            "=IFERROR([@NAV]/[@Peak_NAV]-1,0)",
        ])

    fmt = {2: "#,##0.00", 3: "#,##0.00", 4: "#,##0.00",
           5: "0.00%", 6: "0.00%", 7: "#,##0.00", 8: "0.00%"}
    _write_table(ws, "tbl_EquityCurve", headers, rows, number_formats=fmt)

    if rows:
        _apply_pos_neg_cf(ws, ["E", "F", "H"], len(rows) + 1)


def _write_sector_exposure(wb, positions: pd.DataFrame) -> None:
    ws = wb["Sector_Exposure"]
    _clear_sheet(ws)

    headers = ["Sector", "Market_Value", "Weight"]
    rows = []
    if not positions.empty:
        sectors = sorted(positions["Sector"].dropna().unique())
        sectors = [s for s in sectors if s]
        for sector in sectors:
            rows.append([
                sector,
                '=SUMIFS(tbl_Positions[Market_Value],tbl_Positions[Sector],[@Sector])',
                "=IFERROR([@Market_Value]/SUM(tbl_SectorExposure[Market_Value]),0)",
            ])

    fmt = {2: "#,##0.00", 3: "0.00%"}
    _write_table(ws, "tbl_SectorExposure", headers, rows, number_formats=fmt)


def _write_geo_exposure(wb, positions: pd.DataFrame) -> None:
    ws = wb["Geographic_Exposure"]
    _clear_sheet(ws)

    headers = ["Region", "Market_Value", "Weight"]
    rows = []
    if not positions.empty:
        regions = sorted(positions["Region"].dropna().unique())
        regions = [r for r in regions if r]
        for region in regions:
            rows.append([
                region,
                '=SUMIFS(tbl_Positions[Market_Value],tbl_Positions[Region],[@Region])',
                "=IFERROR([@Market_Value]/SUM(tbl_GeoExposure[Market_Value]),0)",
            ])

    fmt = {2: "#,##0.00", 3: "0.00%"}
    _write_table(ws, "tbl_GeoExposure", headers, rows, number_formats=fmt)


def _write_dashboard(wb, eq_curve: pd.DataFrame, positions: pd.DataFrame,
                     market: pd.DataFrame) -> None:
    ws = wb["Dashboard"]
    _clear_sheet(ws, max_col_override=30)
    ws._charts.clear()

    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 4

    ws.merge_cells("B2:E2")
    ws["B2"] = "Portfolio Dashboard"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = Alignment(horizontal="left")

    kpis = [
        ("NAV", "=IFERROR(TAKE(tbl_EquityCurve[NAV],-1),0)", "#,##0.00"),
        ("Total Return", "=IFERROR(TAKE(tbl_EquityCurve[Cumulative_Return],-1),0)", "0.00%"),
        ("Daily Return", "=IFERROR(TAKE(tbl_EquityCurve[Daily_Return],-1),0)", "0.00%"),
        (
            "Sharpe Ratio",
            "=IFERROR((AVERAGE(tbl_EquityCurve[Daily_Return])*252-RiskFreeRate)"
            "/(STDEV(tbl_EquityCurve[Daily_Return])*SQRT(252)),0)",
            "0.00",
        ),
        ("Max Drawdown", "=IFERROR(MIN(tbl_EquityCurve[Drawdown]),0)", "0.00%"),
        ("VaR (95% Daily)", "=IFERROR(PERCENTILE(tbl_EquityCurve[Daily_Return],0.05),0)", "0.00%"),
        ('# Positions', '=COUNTIF(tbl_Positions[Units],">"&0)', "0"),
        ("Start Date", "=StartDate", "@"),
    ]

    for idx, (label, formula, fmt) in enumerate(kpis):
        row = 4 + idx
        ws.cell(row=row, column=2, value=label).font = KPI_LABEL_FONT
        val_cell = ws.cell(row=row, column=3, value=formula)
        val_cell.font = KPI_VALUE_FONT
        val_cell.number_format = fmt

    _write_correlation_matrix(ws, market, start_col=5, start_row=4)
    _write_nav_chart(wb, ws, eq_curve)
    _write_allocation_chart(wb, ws, positions)


def _write_correlation_matrix(ws, market: pd.DataFrame, start_col: int, start_row: int) -> None:
    ws.cell(row=start_row - 1, column=start_col, value="Correlation Matrix").font = KPI_LABEL_FONT

    if market.empty:
        return

    wide = market.pivot_table(index="Date", columns="Ticker", values="Close")
    if wide.shape[1] < 2:
        return

    returns = wide.pct_change().dropna()
    if returns.empty:
        return

    corr = returns.corr()
    tickers = list(corr.columns)

    for ci, t in enumerate(tickers):
        cell = ws.cell(row=start_row, column=start_col + 1 + ci, value=t)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for ri, t_row in enumerate(tickers):
        label_cell = ws.cell(row=start_row + 1 + ri, column=start_col, value=t_row)
        label_cell.font = Font(name="Segoe UI", bold=True, size=10)
        for ci, t_col in enumerate(tickers):
            val = corr.loc[t_row, t_col]
            cell = ws.cell(row=start_row + 1 + ri, column=start_col + 1 + ci, value=round(float(val), 2))
            cell.number_format = "0.00"

    data_start = f"{get_column_letter(start_col + 1)}{start_row + 1}"
    data_end = f"{get_column_letter(start_col + len(tickers))}{start_row + len(tickers)}"
    rng = f"{data_start}:{data_end}"
    ws.conditional_formatting.add(
        rng,
        ColorScaleRule(
            start_type="num", start_value=-1, start_color="B81D13",
            mid_type="num", mid_value=0, mid_color="FFFFFF",
            end_type="num", end_value=1, end_color="007A33",
        ),
    )

    for ci in range(start_col, start_col + len(tickers) + 1):
        ws.column_dimensions[get_column_letter(ci)].width = max(
            ws.column_dimensions[get_column_letter(ci)].width or 0, 10,
        )


def _write_nav_chart(wb, ws, eq_curve: pd.DataFrame) -> None:
    if eq_curve.empty:
        return

    eq_ws = wb["Equity_Curve"]
    num_rows = len(eq_curve)

    chart = LineChart()
    chart.title = "Portfolio NAV"
    chart.style = 10
    chart.y_axis.title = "NAV ($)"
    chart.x_axis.title = "Date"
    chart.width = 32
    chart.height = 14
    chart.y_axis.numFmt = "#,##0"

    data_ref = ChartRef(eq_ws, min_col=4, min_row=1, max_row=num_rows + 1)
    cats_ref = ChartRef(eq_ws, min_col=1, min_row=2, max_row=num_rows + 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.legend = None

    series = chart.series[0]
    series.graphicalProperties.line.width = 22000

    ws.add_chart(chart, "B14")


def _write_allocation_chart(wb, ws, positions: pd.DataFrame) -> None:
    if positions.empty:
        return

    pos_ws = wb["Positions"]
    num_rows = len(positions)

    chart = PieChart()
    chart.title = "Portfolio Allocation"
    chart.style = 10
    chart.width = 18
    chart.height = 14

    data_ref = ChartRef(pos_ws, min_col=10, min_row=1, max_row=num_rows + 1)
    cats_ref = ChartRef(pos_ws, min_col=1, min_row=2, max_row=num_rows + 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showVal = False

    ws.add_chart(chart, "B30")


def main() -> None:
    args = parse_args()
    workbook_path = Path(args.workbook)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(workbook_path)

    start_capital = float(_read_setting(wb, "Start Capital", 100000))

    raw_trades = _read_table_df(wb, "tbl_TradeLog")
    raw_market = _read_table_df(wb, "tbl_MarketData")
    raw_reference = _read_table_df(wb, "tbl_ReferenceData")

    trades = _prepare_trades(raw_trades)
    market = _prepare_market(raw_market)
    reference = _prepare_reference(raw_reference)

    if trades.empty:
        print("Trade_Log is empty, writing empty analytics sheets")

    trading_dates = sorted(market["Date"].unique()) if not market.empty else []

    daily_units = compute_daily_units(trades, trading_dates)
    daily_cash = compute_daily_cash(trades, trading_dates, start_capital)
    positions = compute_positions(trades, market, reference)
    eq_curve = compute_equity_curve(daily_units, daily_cash, market)

    _write_daily_units(wb, daily_units)
    _write_daily_cash(wb, daily_cash)
    _write_positions(wb, positions)
    _write_equity_curve(wb, eq_curve)
    _write_sector_exposure(wb, positions)
    _write_geo_exposure(wb, positions)
    _write_dashboard(wb, eq_curve, positions, market)

    wb.save(workbook_path)
    n_dates = len(trading_dates)
    n_tickers = len(positions) if not positions.empty else 0
    print(f"Analytics rebuilt: {n_tickers} positions, {n_dates} trading days -> {workbook_path}")


if __name__ == "__main__":
    main()
