"""
Mandiri Excel Statement Parser

Structure (from inspection):
    - Header row at Excel row 16 (index 15): No | Tanggal | Keterangan | Dana Masuk (IDR) | Dana Keluar (IDR) | Saldo (IDR)
    - Columns (0-indexed): 0=No, 4=Tanggal, 7=Keterangan, 15=Dana Masuk, 18=Dana Keluar, 21=Saldo
    - Transaction rows: integer in col 0 identifies a transaction row
    - Odd rows after each txn row contain the time (e.g. '23:59:59 WIB') — skipped
    - Date format: '01 Jan 2026'
    - Number format: '3.500,00' (dot=thousands, comma=decimal) — stored as strings
    - Dana Masuk = incoming (credit), Dana Keluar = outgoing (debit)
    - Keterangan may contain embedded newlines with recipient/reference details
"""

import re
import warnings
from datetime import datetime, date
import pandas as pd


MONTH_MAP = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5,  "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
    "Mei": 5, "Agu": 8, "Okt": 10, "Des": 12,
}

COL_NO       = 0
COL_TANGGAL  = 4
COL_KETER    = 7
COL_MASUK    = 15   # Dana Masuk / incoming / credit
COL_KELUAR   = 18   # Dana Keluar / outgoing / debit
COL_SALDO    = 21

HEADER_ROW   = 15   # 0-indexed row where 'No', 'Tanggal', etc appear


def _parse_date(s):
    if not isinstance(s, str):
        return None
    s = s.strip().split("\n")[0]  # take first line if multiline
    m = re.match(r"(\d{1,2})\s+(\w+)\s+(\d{4})", s)
    if m:
        day   = int(m.group(1))
        month = MONTH_MAP.get(m.group(2))
        yr    = int(m.group(3))
        if month:
            return date(yr, month, day)
    return None


def _parse_number(val):
    """'3.500,00' → 3500.0  (dot=thousands, comma=decimal)"""
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        val = val.strip()
        if not val or val.lower() == 'nan':
            return None
        val = val.replace(".", "").replace(",", ".")
        try:
            return float(val)
        except ValueError:
            return None
    return None


def parse_mandiri_excel(filepath):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        df = pd.read_excel(filepath, header=None, dtype=str)

    transactions = []

    for _, row in df.iterrows():
        # A transaction row has an integer in COL_NO
        no_val = row.iloc[COL_NO]
        if pd.isna(no_val):
            continue
        try:
            int(float(no_val))
        except (ValueError, TypeError):
            continue

        date_val   = _parse_date(row.iloc[COL_TANGGAL] if not pd.isna(row.iloc[COL_TANGGAL]) else "")
        keterangan = str(row.iloc[COL_KETER]).strip() if not pd.isna(row.iloc[COL_KETER]) else ""
        # Normalise embedded newlines to spaces
        keterangan = " ".join(keterangan.splitlines())

        masuk  = _parse_number(row.iloc[COL_MASUK]  if not pd.isna(row.iloc[COL_MASUK])  else None)
        keluar = _parse_number(row.iloc[COL_KELUAR] if not pd.isna(row.iloc[COL_KELUAR]) else None)
        saldo  = _parse_number(row.iloc[COL_SALDO]  if not pd.isna(row.iloc[COL_SALDO])  else None)

        if date_val is None:
            continue

        if masuk and masuk > 0:
            txn_type = "Credit"
            debit    = None
            credit   = masuk
        elif keluar and keluar > 0:
            txn_type = "Debit"
            debit    = keluar
            credit   = None
        else:
            continue  # no amount — skip

        transactions.append({
            "date":        date_val,
            "bank":        "Mandiri",
            "description": keterangan,
            "type":        txn_type,
            "debit":       debit,
            "credit":      credit,
            "balance":     saldo,
        })

    return transactions