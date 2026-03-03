"""
Excel Diary Writer

Sheet 1 — Transactions
  Columns: No | Date | Bank | Type | Category | Description | Debit | Credit | Balance
  - Category column has data validation dropdown with full category list
  - Bank rows colour-coded (BCA=blue tint, Mandiri=yellow tint)

Sheet 2 — Monthly Summary
  - Per-bank subtotals
  - Combined totals row
  - Income vs expense ratio
  - Category breakdown (populated once user fills in categories)
  - Two embedded charts: combined balance + per-bank balance
"""

from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, Reference

# ── Palette ──────────────────────────────────────────────────────────────────
C_DARK_BLUE   = "1F4E79"
C_MID_BLUE    = "2E75B6"
C_BCA_ROW     = "DDEEFF"
C_MANDIRI_ROW = "FFF3CD"
C_ALT_ROW     = "F7F7F7"
C_TOTAL_BG    = "D6E4F0"
C_SUBTOTAL_BG = "EBF3FA"
C_WHITE       = "FFFFFF"
C_DEBIT       = "C00000"
C_CREDIT      = "375623"
C_NEUTRAL     = "000000"

THIN  = Side(style="thin",   color="CCCCCC")
THICK = Side(style="medium", color="AAAAAA")
BORDER      = Border(left=THIN,  right=THIN,  top=THIN,  bottom=THIN)
THICK_BOT   = Border(left=THIN,  right=THIN,  top=THIN,  bottom=THICK)

IDR_FMT  = '#,##0'
IDR_FMT0 = '#,##0;(#,##0);"-"'
PCT_FMT  = '0.0%'
DATE_FMT = 'DD/MM/YYYY'

MONTH_NAMES = {
    1:"January",2:"February",3:"March",4:"April",
    5:"May",6:"June",7:"July",8:"August",
    9:"September",10:"October",11:"November",12:"December",
}

# ── Category list (6 main categories) ────────────────────────────────────────
CATEGORIES = [
    "Income",
    "Fixed Expenses",
    "Variable Expenses",
    "Financial Allocation",
    "Lifestyle & Discretionary",
    "Transfers",
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def _font(bold=False, color=C_NEUTRAL, size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def _left(wrap=True):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

def _right():
    return Alignment(horizontal="right", vertical="center")

def _set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def _hdr_cell(ws, row, col, value, bg=C_DARK_BLUE, fg=C_WHITE, size=10, center=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font    = _font(bold=True, color=fg, size=size)
    c.fill    = _fill(bg)
    c.border  = BORDER
    c.alignment = _center() if center else _left(wrap=False)
    return c

def _data_cell(ws, row, col, value, bg, fmt=None, bold=False,
               color=C_NEUTRAL, align="left"):
    c = ws.cell(row=row, column=col, value=value)
    c.font   = _font(bold=bold, color=color)
    c.fill   = _fill(bg)
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    if align == "center":
        c.alignment = _center()
    elif align == "right":
        c.alignment = _right()
    else:
        c.alignment = _left()
    return c


# ── Sheet 1: Transactions ─────────────────────────────────────────────────────
#  Col: A=No B=Date C=Bank D=Type E=Category F=Description G=Debit H=Credit I=Balance

TXN_HEADERS = [
    "No", "Date", "Bank", "Type", "Category",
    "Description", "Debit (IDR)", "Credit (IDR)", "Balance (IDR)"
]
TXN_WIDTHS = {
    "A": 6, "B": 13, "C": 10, "D": 10, "E": 28,
    "F": 50, "G": 18, "H": 18, "I": 18,
}

# Hidden sheet to hold category list for data validation
CAT_SHEET = "_Categories"


def _write_transactions_sheet(wb, ws, transactions):
    ws.title     = "Transactions"
    ws.freeze_panes = "A2"

    # Build hidden category sheet for dropdown source
    cat_ws = wb.create_sheet(CAT_SHEET)
    cat_ws.sheet_state = "hidden"
    for i, cat in enumerate(CATEGORIES, start=1):
        cat_ws.cell(row=i, column=1, value=cat)
    cat_range = f"'{CAT_SHEET}'!$A$1:$A${len(CATEGORIES)}"

    # Header
    for ci, hdr in enumerate(TXN_HEADERS, start=1):
        _hdr_cell(ws, 1, ci, hdr)
    ws.row_dimensions[1].height = 18

    # Data rows
    for ri, txn in enumerate(transactions, start=2):
        bank = txn.get("bank", "")
        bg   = C_BCA_ROW if bank == "BCA" else (C_MANDIRI_ROW if bank == "Mandiri" else C_ALT_ROW)
        txn_type = txn.get("type", "")

        _data_cell(ws, ri, 1,  ri - 1,              bg, align="center")
        _data_cell(ws, ri, 2,  txn.get("date"),     bg, fmt=DATE_FMT, align="center")
        _data_cell(ws, ri, 3,  bank,                bg, bold=True, align="center")
        _data_cell(ws, ri, 4,  txn_type,            bg, align="center",
                   color=C_DEBIT if txn_type=="Debit" else (C_CREDIT if txn_type=="Credit" else C_NEUTRAL))
        _data_cell(ws, ri, 5,  "",                  bg, align="left")   # Category — user fills
        _data_cell(ws, ri, 6,  txn.get("description",""), bg, align="left")
        _data_cell(ws, ri, 7,  txn.get("debit"),    bg, fmt=IDR_FMT0, align="right",
                   color=C_DEBIT if txn.get("debit") else C_NEUTRAL)
        _data_cell(ws, ri, 8,  txn.get("credit"),   bg, fmt=IDR_FMT0, align="right",
                   color=C_CREDIT if txn.get("credit") else C_NEUTRAL)
        _data_cell(ws, ri, 9,  txn.get("balance"),  bg, fmt=IDR_FMT0, align="right")

    # Totals row
    last  = len(transactions) + 1
    total = last + 1
    for ci in range(1, 10):
        bg = C_DARK_BLUE
        c  = ws.cell(row=total, column=ci)
        c.fill   = _fill(bg)
        c.border = THICK_BOT
        c.font   = _font(bold=True, color=C_WHITE)
    ws.cell(row=total, column=6, value="TOTAL").alignment = _right()
    for ci, col in [(7, "G"), (8, "H")]:
        c = ws.cell(row=total, column=ci,
                    value=f"=SUM({col}2:{col}{last})")
        c.number_format = IDR_FMT
        c.alignment     = _right()

    # Category dropdown on every data row
    dv = DataValidation(
        type="list",
        formula1=cat_range,
        allow_blank=True,
        showErrorMessage=False,
    )
    dv.sqref = f"E2:E{last}"
    ws.add_data_validation(dv)

    _set_widths(ws, TXN_WIDTHS)


# ── Sheet 2: Monthly Summary ──────────────────────────────────────────────────

SUMM_WIDTHS = {
    "A": 28, "B": 10, "C": 18, "D": 18, "E": 18, "F": 16, "G": 22,
    # Chart data table columns
    "J": 16, "K": 20, "L": 20, "M": 20,
}

def _write_summary_sheet(wb, ws, transactions):
    ws.title = "Monthly Summary"

    # ── Aggregate ─────────────────────────────────────────────────────────────
    # { (year, month, bank) -> {count, debit, credit, last_balance, last_date} }
    agg = defaultdict(lambda: {
        "count": 0, "debit": 0.0, "credit": 0.0,
        "last_balance": None, "last_date": None
    })
    for txn in transactions:
        d = txn.get("date")
        if not d:
            continue
        key = (d.year, d.month, txn.get("bank", "Unknown"))
        g   = agg[key]
        g["count"]  += 1
        g["debit"]  += txn.get("debit")  or 0.0
        g["credit"] += txn.get("credit") or 0.0
        if txn.get("balance") is not None:
            if g["last_date"] is None or d >= g["last_date"]:
                g["last_date"]    = d
                g["last_balance"] = txn["balance"]

    sorted_keys    = sorted(agg.keys())
    # Get unique (year, month) periods
    periods        = sorted(set((y, m) for y, m, _ in sorted_keys))
    banks          = sorted(set(b for _, _, b in sorted_keys))

    row = 1

    # ── Section 1: Per-Bank Monthly Breakdown ─────────────────────────────────
    _hdr_cell(ws, row, 1, "MONTHLY BREAKDOWN BY BANK",
              bg=C_DARK_BLUE, size=11)
    ws.merge_cells(f"A{row}:G{row}")
    row += 1

    hdrs = ["Period", "Bank", "# Txn", "Total Debit (IDR)",
            "Total Credit (IDR)", "Net Flow (IDR)", "Closing Balance (IDR)"]
    for ci, h in enumerate(hdrs, 1):
        _hdr_cell(ws, row, ci, h, bg=C_MID_BLUE)
    row += 1

    bank_row_start = row

    for (year, month, bank) in sorted_keys:
        g  = agg[(year, month, bank)]
        bg = C_BCA_ROW if bank == "BCA" else (C_MANDIRI_ROW if bank == "Mandiri" else C_ALT_ROW)
        net = g["credit"] - g["debit"]
        _data_cell(ws, row, 1, f"{MONTH_NAMES[month]} {year}", bg, align="center")
        _data_cell(ws, row, 2, bank,                    bg, bold=True, align="center")
        _data_cell(ws, row, 3, g["count"],              bg, align="center")
        _data_cell(ws, row, 4, g["debit"]  or None,     bg, fmt=IDR_FMT0, align="right", color=C_DEBIT)
        _data_cell(ws, row, 5, g["credit"] or None,     bg, fmt=IDR_FMT0, align="right", color=C_CREDIT)
        _data_cell(ws, row, 6, net,                     bg, fmt=IDR_FMT0, align="right",
                   color=C_CREDIT if net >= 0 else C_DEBIT, bold=True)
        _data_cell(ws, row, 7, g["last_balance"],       bg, fmt=IDR_FMT0, align="right")
        row += 1

    bank_row_end = row - 1

    # ── Per-period combined subtotals ─────────────────────────────────────────
    row += 1
    _hdr_cell(ws, row, 1, "COMBINED TOTALS PER PERIOD",
              bg=C_DARK_BLUE, size=11)
    ws.merge_cells(f"A{row}:G{row}")
    row += 1

    hdrs2 = ["Period", "Total Txn", "Total Debit (IDR)", "Total Credit (IDR)",
             "Net Flow (IDR)", "Income/Expense Ratio", ""]
    for ci, h in enumerate(hdrs2, 1):
        _hdr_cell(ws, row, ci, h, bg=C_MID_BLUE)
    row += 1

    combined_row_start = row
    for (year, month) in periods:
        total_debit  = sum(agg[(year, month, b)]["debit"]  for b in banks if (year,month,b) in agg)
        total_credit = sum(agg[(year, month, b)]["credit"] for b in banks if (year,month,b) in agg)
        total_count  = sum(agg[(year, month, b)]["count"]  for b in banks if (year,month,b) in agg)
        net          = total_credit - total_debit
        ratio        = (total_credit / total_debit) if total_debit else None

        _data_cell(ws, row, 1, f"{MONTH_NAMES[month]} {year}", C_SUBTOTAL_BG, align="center", bold=True)
        _data_cell(ws, row, 2, total_count,   C_SUBTOTAL_BG, align="center")
        _data_cell(ws, row, 3, total_debit or None,  C_SUBTOTAL_BG, fmt=IDR_FMT0, align="right", color=C_DEBIT,  bold=True)
        _data_cell(ws, row, 4, total_credit or None, C_SUBTOTAL_BG, fmt=IDR_FMT0, align="right", color=C_CREDIT, bold=True)
        _data_cell(ws, row, 5, net, C_SUBTOTAL_BG, fmt=IDR_FMT0, align="right",
                   color=C_CREDIT if net >= 0 else C_DEBIT, bold=True)
        _data_cell(ws, row, 6, ratio, C_SUBTOTAL_BG, fmt=PCT_FMT, align="center",
                   color=C_CREDIT if (ratio and ratio >= 1) else C_DEBIT)
        row += 1

    combined_row_end = row - 1

    # ── Section 3: Category Summary (formula-driven, updates as user tags) ────
    row += 2
    _hdr_cell(ws, row, 1, "SPENDING BY CATEGORY",
              bg=C_DARK_BLUE, size=11)
    ws.merge_cells(f"A{row}:G{row}")
    row += 1

    note_row = row
    c = ws.cell(row=row, column=1,
                value="⚑  Fill in the Category column in the Transactions sheet — this table updates automatically.")
    c.font      = _font(italic=True, color="595959", size=9)
    c.alignment = _left(wrap=False)
    ws.merge_cells(f"A{row}:G{row}")
    row += 1

    cat_hdrs = ["Category", "# Txn", "Total Debit (IDR)", "Total Credit (IDR)", "Net (IDR)", "", ""]
    for ci, h in enumerate(cat_hdrs, 1):
        _hdr_cell(ws, row, ci, h, bg=C_MID_BLUE)
    row += 1

    cat_section_start = row
    for cat in CATEGORIES:
        bg = C_ALT_ROW if (row % 2 == 0) else C_WHITE
        _data_cell(ws, row, 1, cat, bg, align="left")
        # COUNTIFS / SUMIFS pulling from Transactions sheet
        _data_cell(ws, row, 2,
            f'=COUNTIFS(Transactions!E:E,A{row})',
            bg, align="center")
        _data_cell(ws, row, 3,
            f'=SUMIFS(Transactions!G:G,Transactions!E:E,A{row})',
            bg, fmt=IDR_FMT0, align="right", color=C_DEBIT)
        _data_cell(ws, row, 4,
            f'=SUMIFS(Transactions!H:H,Transactions!E:E,A{row})',
            bg, fmt=IDR_FMT0, align="right", color=C_CREDIT)
        _data_cell(ws, row, 5,
            f'=D{row}-C{row}',
            bg, fmt=IDR_FMT0, align="right")
        # Colour net via conditional — just set neutral; user will see values
        row += 1

    cat_section_end = row - 1

    # Totals row for category section
    _data_cell(ws, row, 1, "TOTAL", C_DARK_BLUE, bold=True, color=C_WHITE, align="right")
    for ci, col in [(2,"B"),(3,"C"),(4,"D"),(5,"E")]:
        c = ws.cell(row=row, column=ci,
                    value=f"=SUM({col}{cat_section_start}:{col}{cat_section_end})")
        c.fill          = _fill(C_DARK_BLUE)
        c.font          = _font(bold=True, color=C_WHITE)
        c.border        = THICK_BOT
        c.number_format = IDR_FMT
        c.alignment     = _right() if ci > 1 else _center()
    row += 2

    _set_widths(ws, SUMM_WIDTHS)

    # ── Charts ────────────────────────────────────────────────────────────────
    # Build a small helper table of dates + per-bank balances for charting.
    # We need a data table on the sheet; put it far to the right (col J onwards).
    chart_data_col_start = 10  # col J

    # Collect all unique dates in sorted order
    all_dates = sorted(set(
        txn["date"] for txn in transactions if txn.get("date") and txn.get("balance") is not None
    ))

    # For each date take the last known balance per bank
    # Build: date | BCA balance | Mandiri balance | combined (sum of last known)
    balance_by_date_bank = {}  # (date, bank) -> balance
    for txn in sorted(transactions, key=lambda t: (t.get("date") or "", t.get("bank",""))):
        if txn.get("date") and txn.get("balance") is not None:
            balance_by_date_bank[(txn["date"], txn["bank"])] = txn["balance"]

    # Section header spanning J-M
    _hdr_cell(ws, 1, chart_data_col_start, "CHART DATA", bg=C_DARK_BLUE, size=11)
    ws.merge_cells(
        start_row=1, start_column=chart_data_col_start,
        end_row=1,   end_column=chart_data_col_start + 3
    )

    # Column headers on row 2
    chart_hdr_row  = 2
    chart_ws_row   = 3   # data starts on row 3

    for ci, label in enumerate(["Date", "BCA Balance", "Mandiri Balance", "Combined Balance"]):
        _hdr_cell(ws, chart_hdr_row, chart_data_col_start + ci, label, bg=C_MID_BLUE)

    # Forward-fill: for each date, carry forward last known balance per bank
    chart_dates = sorted(set(d for d, _ in balance_by_date_bank))

    last_bal = {"BCA": None, "Mandiri": None}
    for d in chart_dates:
        for bank in ["BCA", "Mandiri"]:
            if (d, bank) in balance_by_date_bank:
                last_bal[bank] = balance_by_date_bank[(d, bank)]
        bca_b  = last_bal["BCA"]
        man_b  = last_bal["Mandiri"]
        comb_b = (bca_b or 0) + (man_b or 0) if (bca_b is not None or man_b is not None) else None
        ws.cell(row=chart_ws_row, column=chart_data_col_start,   value=d).number_format = DATE_FMT
        ws.cell(row=chart_ws_row, column=chart_data_col_start+1, value=bca_b)
        ws.cell(row=chart_ws_row, column=chart_data_col_start+2, value=man_b)
        ws.cell(row=chart_ws_row, column=chart_data_col_start+3, value=comb_b)
        chart_ws_row += 1

    chart_data_rows = chart_ws_row - 1  # last data row

    def _make_line_chart(title, series_cols, legend_labels, anchor):
        chart = LineChart()
        chart.title        = title
        chart.style        = 10
        chart.y_axis.title = "Balance (IDR)"
        chart.x_axis.title = "Date"
        chart.height       = 12
        chart.width        = 22
        chart.y_axis.numFmt = "#,##0"

        # Date categories (data rows only, no header)
        dates = Reference(ws,
                          min_col=chart_data_col_start,
                          min_row=3,
                          max_row=chart_data_rows)

        for col_offset in series_cols:
            # Include column header row (min_row=2) so add_data picks up the label
            data_ref = Reference(ws,
                                 min_col=chart_data_col_start + col_offset,
                                 min_row=2,
                                 max_row=chart_data_rows)
            chart.add_data(data_ref, titles_from_data=True)

        chart.set_categories(dates)
        ws.add_chart(chart, anchor)

    # Charts stacked vertically with a 3-row gap between them
    _make_line_chart(
        title="Combined Balance Over Time",
        series_cols=[3],
        legend_labels=["Combined (BCA + Mandiri)"],
        anchor=f"A{row}"
    )
    _make_line_chart(
        title="Balance by Bank Over Time",
        series_cols=[1, 2],
        legend_labels=["BCA", "Mandiri"],
        anchor=f"A{row + 25}"
    )


# ── Entry point ───────────────────────────────────────────────────────────────

def write_diary(transactions, output_path):
    wb = Workbook()

    ws_txn = wb.active
    _write_transactions_sheet(wb, ws_txn, transactions)

    ws_summ = wb.create_sheet()
    _write_summary_sheet(wb, ws_summ, transactions)

    wb.save(output_path)