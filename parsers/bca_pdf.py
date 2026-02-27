"""
BCA PDF Statement Parser

Layout (from coordinate analysis):
    - TANGGAL      x0 ≈ 33-85
    - KETERANGAN   x0 ≈ 86-305
    - CBG          x0 ≈ 306-375
    - MUTASI       x0 ≈ 376-498  — number like '175,272.00' followed by 'DB' or 'CR'
    - SALDO        x0 ≈ 499+     — number like '2,731,197.35'

Number format: comma = thousands separator, dot = decimal  e.g. 1,993,571.32
Debit/Credit: determined by the 'DB'/'CR' suffix word in the MUTASI column, confirmed by keywords in KETERANGAN.
"""

import re
from datetime import datetime
import pdfplumber


COL_DATE_MAX      = 85
COL_KETERANGAN_MIN = 86
COL_KETERANGAN_MAX = 305
COL_CBG_MIN       = 306
COL_CBG_MAX       = 375
COL_MUTASI_MIN    = 376
COL_MUTASI_MAX    = 498
COL_SALDO_MIN     = 499

DATE_RE   = re.compile(r"^\d{2}/\d{2}$")
# BCA: comma=thousands, dot=decimal  e.g. '175,272.00' or '1,993,571.32'
NUMBER_RE = re.compile(r"^[\d,]+\.\d{2}$")

MONTH_MAP = {
    "JANUARI": 1, "FEBRUARI": 2, "MARET": 3, "APRIL": 4,
    "MEI": 5, "JUNI": 6, "JULI": 7, "AGUSTUS": 8,
    "SEPTEMBER": 9, "OKTOBER": 10, "NOVEMBER": 11, "DESEMBER": 12,
}


def _classify_word(w):
    x0 = w["x0"]
    if x0 <= COL_DATE_MAX:
        return "date"
    elif COL_KETERANGAN_MIN <= x0 <= COL_KETERANGAN_MAX:
        return "keterangan"
    elif COL_CBG_MIN <= x0 <= COL_CBG_MAX:
        return "cbg"
    elif COL_MUTASI_MIN <= x0 <= COL_MUTASI_MAX:
        return "mutasi"
    elif x0 >= COL_SALDO_MIN:
        return "saldo"
    return "unknown"


def _parse_bca_number(s):
    """
    BCA format: '175,272.00'  (comma=thousands, dot=decimal)
    Strip any trailing DB/CR suffix first.
    """
    s = s.strip()
    # Remove trailing DB / CR suffix if accidentally included
    s = re.sub(r'\s*(DB|CR)\s*$', '', s, flags=re.IGNORECASE).strip()
    # Remove thousands commas, keep decimal dot
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def _determine_type_from_mutasi_words(mutasi_words):
    """
    The MUTASI column has two words per transaction:
        word 1: the number  e.g. '175,272.00'
        word 2: the suffix  'DB' or 'CR'
    """
    for w in mutasi_words:
        t = w["text"].upper()
        if t == "DB":
            return "Debit"
        if t == "CR":
            return "Credit"
    return None


def _determine_type_from_keterangan(keterangan):
    k = keterangan.upper()
    if " DB" in k or k.startswith("DB ") or "DEBIT" in k or "TRANSAKSI DEBIT" in k:
        return "Debit"
    if " CR" in k or k.startswith("CR ") or "KREDIT" in k or "CREDIT" in k or "BI-FAST CR" in k:
        return "Credit"
    return "Unknown"


def parse_bca_pdf(filepath):
    transactions = []
    year = datetime.now().year

    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            full_text = (page.extract_text() or "").upper()

            # Detect period year/month from header
            for line in full_text.split("\n"):
                if "PERIODE" in line:
                    for mname, mnum in MONTH_MAP.items():
                        if mname in line:
                            pass  # month not needed for date parsing since DD/MM format
                    m = re.search(r"\b(20\d{2})\b", line)
                    if m:
                        year = int(m.group(1))

            words = page.extract_words(keep_blank_chars=False)

            # Group words by row (3px tolerance)
            rows = {}
            for w in words:
                row_key = round(w["top"] / 3) * 3
                rows.setdefault(row_key, []).append(w)

            sorted_rows = sorted(rows.items())
            in_transaction_area = False

            current_date = None
            current_keterangan_parts = []
            current_mutasi = None
            current_mutasi_type = None  # 'Debit' or 'Credit' from DB/CR suffix
            current_saldo = None

            def flush_transaction():
                nonlocal current_date, current_keterangan_parts, current_mutasi
                nonlocal current_mutasi_type, current_saldo

                if current_date is None or not current_keterangan_parts:
                    current_date = None
                    current_keterangan_parts = []
                    current_mutasi = None
                    current_mutasi_type = None
                    current_saldo = None
                    return

                keterangan = " ".join(current_keterangan_parts)
                ku = keterangan.upper()

                if ku.startswith("SALDO") or ku in {"TANGGAL", "KETERANGAN", "CBG", "MUTASI", "SALDO"}:
                    current_date = None
                    current_keterangan_parts = []
                    current_mutasi = None
                    current_mutasi_type = None
                    current_saldo = None
                    return

                # Prefer DB/CR suffix from the mutasi column; fall back to keterangan keywords
                txn_type = current_mutasi_type or _determine_type_from_keterangan(keterangan)
                debit  = current_mutasi if txn_type == "Debit"  else None
                credit = current_mutasi if txn_type == "Credit" else None

                transactions.append({
                    "date":        current_date,
                    "bank":        "BCA",
                    "description": keterangan,
                    "type":        txn_type,
                    "debit":       debit,
                    "credit":      credit,
                    "balance":     current_saldo,
                })
                current_date = None
                current_keterangan_parts = []
                current_mutasi = None
                current_mutasi_type = None
                current_saldo = None

            for _, row_words in sorted_rows:
                row_words_sorted = sorted(row_words, key=lambda w: w["x0"])
                row_text = " ".join(w["text"] for w in row_words_sorted).upper().strip()

                if "TANGGAL" in row_text and "KETERANGAN" in row_text and "MUTASI" in row_text:
                    in_transaction_area = True
                    continue

                if not in_transaction_area:
                    continue

                if any(p in row_text for p in ["BERSAMBUNG", "HALAMAN BERIKUT", "CATATAN"]):
                    continue

                if "SALDO AKHIR" in row_text:
                    flush_transaction()
                    continue

                date_words   = [w for w in row_words_sorted if _classify_word(w) == "date"]
                keter_words  = [w for w in row_words_sorted if _classify_word(w) == "keterangan"]
                mutasi_words = [w for w in row_words_sorted if _classify_word(w) == "mutasi"]
                saldo_words  = [w for w in row_words_sorted if _classify_word(w) == "saldo"]

                date_str = date_words[0]["text"] if date_words else None

                if date_str and DATE_RE.match(date_str):
                    flush_transaction()
                    day, mon = date_str.split("/")
                    try:
                        current_date = datetime(year, int(mon), int(day)).date()
                    except ValueError:
                        current_date = None

                if keter_words:
                    keter_text = " ".join(w["text"] for w in keter_words).strip()
                    if not keter_text.upper().startswith("TANGGAL"):
                        current_keterangan_parts.append(keter_text)

                # MUTASI column: expect a number word then 'DB' or 'CR'
                if mutasi_words:
                    # Try to get DB/CR type from suffix word
                    t = _determine_type_from_mutasi_words(mutasi_words)
                    if t:
                        current_mutasi_type = t
                    # Parse the number (first word that matches number pattern)
                    for w in mutasi_words:
                        if NUMBER_RE.match(w["text"]):
                            val = _parse_bca_number(w["text"])
                            if val is not None:
                                current_mutasi = val
                            break

                if saldo_words:
                    # Saldo is right-aligned; take the last word that looks like a number
                    for w in reversed(saldo_words):
                        if NUMBER_RE.match(w["text"]):
                            val = _parse_bca_number(w["text"])
                            if val is not None:
                                current_saldo = val
                            break

            flush_transaction()

    return transactions