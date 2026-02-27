"""
Excel Diary Writer

Produces a workbook with two sheets:
    1. "Transactions" — all transactions, sorted by date, with full detail
    2. "Monthly Summary" — one row per bank per month with totals and net flow
"""

from collections import defaultdict
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter


# ── Colours ──────────────────────────────────────────────────────────────────
CLR_HEADER_BG   = "1F4E79"   # Dark blue
CLR_HEADER_FONT = "FFFFFF"   # White
CLR_BCA_BG      = "DDEEFF"   # Light blue (BCA row tint)
CLR_MANDIRI_BG  = "FFF3CD"   # Light yellow (Mandiri row tint)
CLR_DEBIT_FONT  = "C00000"   # Dark red for debit amounts
CLR_CREDIT_FONT = "375623"   # Dark green for credit amounts
CLR_SUMHDR_BG   = "2E75B6"   # Summary sheet header
CLR_SUBTOTAL_BG = "BDD7EE"   # Summary subtotal row
CLR_ALT_ROW     = "F2F2F2"   # Alternate row grey

IDR_FORMAT = '#,##0'
DATE_FORMAT = 'DD/MM/YYYY'

THIN = Side(style="thin", color="AAAAAA")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _hdr_font(size=10):
    return Font(name="Arial", bold=True, color=CLR_HEADER_FONT, size=size)


def _body_font(bold=False, color="000000", size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)


def _fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)


def _center():
    return Alignment(horizontal="center", vertical="center", wrap_text=False)


def _left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def _right():
    return Alignment(horizontal="right", vertical="center")


def _set_col_widths(ws, widths):
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


# ── Sheet 1: Transactions ─────────────────────────────────────────────────────

TXN_HEADERS = [
    "No", "Date", "Bank", "Type", "Description", "Debit (IDR)", "Credit (IDR)", "Balance (IDR)"
]
TXN_COL_WIDTHS = {
    "A": 6, "B": 13, "C": 10, "D": 10, "E": 48, "F": 18, "G": 18, "H": 18
}


def _write_transactions_sheet(ws, transactions):
    ws.title = "Transactions"
    ws.freeze_panes = "A2"

    # Header row
    for col_idx, hdr in enumerate(TXN_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=hdr)
        cell.font = _hdr_font()
        cell.fill = _fill(CLR_HEADER_BG)
        cell.alignment = _center()
        cell.border = THIN_BORDER

    # Data rows
    for row_idx, txn in enumerate(transactions, start=2):
        bank = txn.get("bank", "")
        row_bg = CLR_BCA_BG if bank == "BCA" else (CLR_MANDIRI_BG if bank == "Mandiri" else CLR_ALT_ROW)

        values = [
            row_idx - 1,
            txn.get("date"),
            bank,
            txn.get("type", ""),
            txn.get("description", ""),
            txn.get("debit"),
            txn.get("credit"),
            txn.get("balance"),
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = _fill(row_bg)
            cell.border = THIN_BORDER

            if col_idx == 1:  # No
                cell.alignment = _center()
                cell.font = _body_font()
            elif col_idx == 2:  # Date
                cell.number_format = DATE_FORMAT
                cell.alignment = _center()
                cell.font = _body_font()
            elif col_idx == 3:  # Bank
                cell.alignment = _center()
                cell.font = _body_font(bold=True)
            elif col_idx == 4:  # Type
                color = CLR_DEBIT_FONT if val == "Debit" else (CLR_CREDIT_FONT if val == "Credit" else "000000")
                cell.alignment = _center()
                cell.font = _body_font(color=color)
            elif col_idx == 5:  # Description
                cell.alignment = _left()
                cell.font = _body_font()
            elif col_idx == 6:  # Debit
                cell.number_format = IDR_FORMAT
                cell.alignment = _right()
                cell.font = _body_font(color=CLR_DEBIT_FONT if val else "000000")
            elif col_idx == 7:  # Credit
                cell.number_format = IDR_FORMAT
                cell.alignment = _right()
                cell.font = _body_font(color=CLR_CREDIT_FONT if val else "000000")
            elif col_idx == 8:  # Balance
                cell.number_format = IDR_FORMAT
                cell.alignment = _right()
                cell.font = _body_font()

    # Totals row
    total_row = len(transactions) + 2
    ws.cell(row=total_row, column=5, value="TOTAL").font = _hdr_font(10)
    ws.cell(row=total_row, column=5).fill = _fill(CLR_HEADER_BG)
    ws.cell(row=total_row, column=5).alignment = _right()

    last_data = len(transactions) + 1
    for col_idx in [6, 7, 8]:
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=total_row, column=col_idx,
                value=f"=SUM({col_letter}2:{col_letter}{last_data})")
        cell.number_format = IDR_FORMAT
        cell.font = _hdr_font(10)
        cell.fill = _fill(CLR_HEADER_BG)
        cell.alignment = _right()
        cell.border = THIN_BORDER

    # Fix col 5 border too
    for col_idx in range(1, 9):
        c = ws.cell(row=total_row, column=col_idx)
        c.border = THIN_BORDER
        if col_idx < 5:
            c.fill = _fill(CLR_HEADER_BG)

    _set_col_widths(ws, TXN_COL_WIDTHS)
    ws.row_dimensions[1].height = 18


# ── Sheet 2: Monthly Summary ──────────────────────────────────────────────────

SUMM_HEADERS = [
    "Month", "Year", "Bank", "# Transactions", "Total Debit (IDR)", "Total Credit (IDR)",
    "Net Flow (IDR)", "Closing Balance (IDR)"
]
SUMM_COL_WIDTHS = {
    "A": 12, "B": 8, "C": 10, "D": 16, "E": 22, "F": 22, "G": 18, "H": 24
}

MONTH_NAMES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}


def _write_summary_sheet(ws, transactions):
    ws.title = "Monthly Summary"
    ws.freeze_panes = "A2"

    # Header
    for col_idx, hdr in enumerate(SUMM_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=hdr)
        cell.font = _hdr_font()
        cell.fill = _fill(CLR_SUMHDR_BG)
        cell.alignment = _center()
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 18

    # Aggregate: { (year, month, bank) -> {txns, debit, credit, last_balance} }
    agg = defaultdict(lambda: {"count": 0, "debit": 0.0, "credit": 0.0, "last_balance": None, "last_date": None})

    for txn in transactions:
        d = txn.get("date")
        if d is None:
            continue
        key = (d.year, d.month, txn.get("bank", "Unknown"))
        g = agg[key]
        g["count"] += 1
        g["debit"] += txn.get("debit") or 0.0
        g["credit"] += txn.get("credit") or 0.0
        if txn.get("balance") is not None:
            if g["last_date"] is None or d >= g["last_date"]:
                g["last_date"] = d
                g["last_balance"] = txn["balance"]

    sorted_keys = sorted(agg.keys())
    row_idx = 2

    for key in sorted_keys:
        year, month, bank = key
        g = agg[key]
        row_bg = CLR_BCA_BG if bank == "BCA" else (CLR_MANDIRI_BG if bank == "Mandiri" else CLR_ALT_ROW)

        net = g["credit"] - g["debit"]

        row_data = [
            MONTH_NAMES[month],
            year,
            bank,
            g["count"],
            g["debit"] if g["debit"] else None,
            g["credit"] if g["credit"] else None,
            net,
            g["last_balance"],
        ]

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = _fill(row_bg)
            cell.border = THIN_BORDER

            if col_idx in (1, 2, 3):
                cell.alignment = _center()
                cell.font = _body_font(bold=(col_idx == 3))
            elif col_idx == 4:
                cell.alignment = _center()
                cell.font = _body_font()
            elif col_idx in (5, 6, 8):
                cell.number_format = IDR_FORMAT
                cell.alignment = _right()
                color = CLR_DEBIT_FONT if col_idx == 5 else (CLR_CREDIT_FONT if col_idx == 6 else "000000")
                cell.font = _body_font(color=color)
            elif col_idx == 7:
                cell.number_format = IDR_FORMAT
                cell.alignment = _right()
                color = CLR_CREDIT_FONT if net >= 0 else CLR_DEBIT_FONT
                cell.font = _body_font(bold=True, color=color)

        row_idx += 1

    _set_col_widths(ws, SUMM_COL_WIDTHS)


# ── Entry point ───────────────────────────────────────────────────────────────

def write_diary(transactions, output_path):
    wb = Workbook()
    ws_txn = wb.active
    _write_transactions_sheet(ws_txn, transactions)

    ws_summ = wb.create_sheet()
    _write_summary_sheet(ws_summ, transactions)

    wb.save(output_path)
