from __future__ import annotations

from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

DEFAULT_CONFIG_PATH = Path("config/portfolio.yaml")
DEFAULT_OUTPUT_PATH = Path("output/portfolio_workbook.xlsx")

def load_output_path(config_path: Path = DEFAULT_CONFIG_PATH) -> Path:
    if not config_path.exists():
        return DEFAULT_OUTPUT_PATH
    with config_path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}
    paths = raw.get("paths", {})
    return Path(paths.get("workbook_output", str(DEFAULT_OUTPUT_PATH)))


SHEETS = [
    "Settings",
    "Trade_Log",
    "Market_Data",
    "Reference_Data",
    "Daily_Units",
    "Daily_Cash",
    "Positions",
    "Equity_Curve",
    "Dashboard",
    "Sector_Exposure",
    "Geographic_Exposure",
]

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def _autosize_columns(ws) -> None:
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(12, min(32, max_len + 2))


def _style_header_row(ws) -> None:
    for cell in ws[1]:
        if cell.value is None:
            continue
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT


def _make_table(ws, table_name: str, headers: list[str], seed_row: list) -> Table:
    ws.append(headers)
    ws.append(seed_row)
    table_ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    table = Table(displayName=table_name, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    _style_header_row(ws)
    _autosize_columns(ws)
    ws.freeze_panes = "A2"
    return table


def _apply_pos_neg_cf(ws, start_cell: str, end_cell: str) -> None:
    green = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    rng = f"{start_cell}:{end_cell}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=red))


def build_workbook(output_path: Path | None = None) -> Path:
    output_path = output_path or load_output_path()
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name in SHEETS:
        wb.create_sheet(sheet_name)

    settings = wb["Settings"]
    settings["A1"] = "Parameter"
    settings["B1"] = "Value"
    settings["A2"] = "Start Capital"
    settings["B2"] = 100000
    settings["A3"] = "Start Date"
    settings["B3"] = "2024-01-01"
    settings["A4"] = "Risk Free Rate"
    settings["B4"] = 0.02
    settings["A5"] = "Base Currency"
    settings["B5"] = "USD"
    _style_header_row(settings)
    _autosize_columns(settings)
    settings.freeze_panes = "A2"

    wb.defined_names.add(DefinedName("StartCapital", attr_text="Settings!$B$2"))
    wb.defined_names.add(DefinedName("StartDate", attr_text="Settings!$B$3"))
    wb.defined_names.add(DefinedName("RiskFreeRate", attr_text="Settings!$B$4"))
    wb.defined_names.add(DefinedName("BaseCcy", attr_text="Settings!$B$5"))

    _make_table(
        wb["Trade_Log"],
        "tbl_TradeLog",
        [
            "Date",
            "IB_ExecId",
            "Ticker",
            "Action",
            "Shares",
            "Price",
            "Commission",
            "permId",
            "account",
            "exchange",
            "currency",
            "Gross_Value",
            "Net_Cost",
            "Signed_Shares",
            "Signed_Cash",
            "Cash_Balance",
        ],
        [
            "",
            "",
            "",
            0,
            0,
            0,
            "",
            "",
            "",
            "",
            "=[@Shares]*[@Price]",
            "=[@Gross_Value]+[@Commission]",
            "=IF(UPPER([@Action])=\"BUY\",[@Shares],-[@Shares])",
            "=-[@Signed_Shares]*[@Price]-[@Commission]",
            "=StartCapital+SUM(INDEX(tbl_TradeLog[Signed_Cash],1):[@Signed_Cash])",
        ],
    )

    _make_table(
        wb["Market_Data"],
        "tbl_MarketData",
        ["Date", "Ticker", "Open", "High", "Low", "Close", "Adj_Close", "Volume"],
        ["", "", 0, 0, 0, 0, 0, 0],
    )

    _make_table(
        wb["Reference_Data"],
        "tbl_ReferenceData",
        ["Ticker", "Name", "Sector", "Region", "Currency", "Benchmark"],
        ["", "", "", "", "", ""],
    )

    _make_table(
        wb["Daily_Units"],
        "tbl_DailyUnits",
        ["Date", "Ticker", "Units"],
        [
            "",
            "",
            "=SUMIFS(tbl_TradeLog[Signed_Shares],tbl_TradeLog[Ticker],[@Ticker],tbl_TradeLog[Date],\"<=\"&[@Date])",
        ],
    )

    _make_table(
        wb["Daily_Cash"],
        "tbl_DailyCash",
        ["Date", "Cash_Balance"],
        [
            "",
            "=StartCapital+SUMIFS(tbl_TradeLog[Signed_Cash],tbl_TradeLog[Date],\"<=\"&[@Date])",
        ],
    )

    _make_table(
        wb["Positions"],
        "tbl_Positions",
        ["Ticker", "Units", "Last_Price", "Market_Value", "Cost_Basis", "Unrealized_PnL", "Weight"],
        [
            "",
            "=LET(t,[@Ticker],u,FILTER(tbl_DailyUnits[Units],tbl_DailyUnits[Ticker]=t),IFERROR(TAKE(u,-1),0))",
            "=XLOOKUP([@Ticker],tbl_MarketData[Ticker],tbl_MarketData[Close],0)",
            "=[@Units]*[@Last_Price]",
            "=LET(t,[@Ticker],cost,SUMIFS(tbl_TradeLog[Net_Cost],tbl_TradeLog[Ticker],t),qty,SUMIFS(tbl_TradeLog[Signed_Shares],tbl_TradeLog[Ticker],t),IFERROR(cost/qty,0))",
            "=[@Market_Value]-([@Units]*[@Cost_Basis])",
            "=IFERROR([@Market_Value]/SUM(tbl_Positions[Market_Value]),0)",
        ],
    )

    _make_table(
        wb["Equity_Curve"],
        "tbl_EquityCurve",
        ["Date", "Portfolio_Value", "Cash_Balance", "NAV", "Daily_Return", "Peak_NAV", "Drawdown"],
        [
            "",
            "=SUM(tbl_Positions[Market_Value])",
            "=XLOOKUP([@Date],tbl_DailyCash[Date],tbl_DailyCash[Cash_Balance],0)",
            "=[@Portfolio_Value]+[@Cash_Balance]",
            "=IFERROR([@NAV]/OFFSET([@NAV],-1,0)-1,0)",
            "=MAX(INDEX(tbl_EquityCurve[NAV],1):[@NAV])",
            "=IFERROR([@NAV]/[@Peak_NAV]-1,0)",
        ],
    )

    for ws_name in ["Dashboard", "Sector_Exposure", "Geographic_Exposure"]:
        ws = wb[ws_name]
        ws["A1"] = f"{ws_name} (formula-only analytics)"
        _style_header_row(ws)
        ws.freeze_panes = "A2"
        _autosize_columns(ws)

    _apply_pos_neg_cf(wb["Trade_Log"], "J2", "J2000")
    _apply_pos_neg_cf(wb["Positions"], "F2", "F2000")
    _apply_pos_neg_cf(wb["Equity_Curve"], "E2", "G2000")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    saved = build_workbook()
    print(f"Workbook saved to: {saved}")
