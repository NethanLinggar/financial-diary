# Financial Diary Generator

Parses BCA and Mandiri e-statements and produces a monthly financial diary in Excel.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
# BCA PDF only
python main.py --bca bca.pdf --output diary_jan2026.xlsx

# Mandiri Excel only
python main.py --mandiri mandiri.xlsx --output diary_jan2026.xlsx

# BCA PDF + Mandiri Excel
python main.py --bca bca.pdf --mandiri mandiri.xlsx --output diary_jan2026.xlsx
```

## Output

The generated Excel file has two sheets:

### Sheet 1 — Transactions

| Column        | Description                             |
| ------------- | --------------------------------------- |
| No            | Row number                              |
| Date          | Transaction date (DD/MM/YYYY)           |
| Bank          | BCA or Mandiri                          |
| Type          | Debit or Credit                         |
| Description   | Keterangan / Remarks from the statement |
| Debit (IDR)   | Amount out (red)                        |
| Credit (IDR)  | Amount in (green)                       |
| Balance (IDR) | Running balance after transaction       |

Rows are colour-coded: **blue tint** = BCA, **yellow tint** = Mandiri.

### Sheet 2 — Monthly Summary

One row per bank per month, showing:

- Transaction count
- Total debit, total credit
- Net flow (green if positive, red if negative)
- Closing balance (last known balance in the period)

## Statement Formats Supported

| Bank    | Format           | Notes                                    |
| ------- | ---------------- | ---------------------------------------- |
| BCA     | PDF (text-based) | Rekening Tahapan / Tahapan Xpresi        |
| Mandiri | Excel (.xlsx)    | Auto-detects header row and column names |

## Debit/Credit Detection

- **BCA**: Keywords in KETERANGAN — `DB` / `DEBIT` = debit; `CR` / `KREDIT` = credit
- **Mandiri Excel**: Same as PDF (negative nominal = debit)

## Notes

- If you have statements spanning multiple months, run with all PDFs for the same period;
  the summary sheet will automatically group by month.
- The Mandiri Excel parser auto-detects column positions — if your Excel layout differs,
  check that column headers contain the words `Tanggal`, `Keterangan`, `Nominal`, `Saldo`.
