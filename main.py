"""
Financial Diary Generator
Parses BCA (PDF) and Mandiri (PDF or Excel) e-statements
and produces a monthly financial diary Excel workbook.

Usage:
    python main.py --bca <bca.pdf> --mandiri <mandiri.pdf or mandiri.xlsx> --output <out.xlsx>
"""

import argparse
import getpass
import io
import sys
import tempfile
import os
from parsers.bca_pdf import parse_bca_pdf
from parsers.mandiri_excel import parse_mandiri_excel
from output.excel_writer import write_diary


def _decrypt_pdf(path):
    """Prompt for password and return a decrypted copy as a temp file path."""
    try:
        import pikepdf
    except ImportError:
        print("Error: pikepdf is required for password-protected PDFs. Run: pip install pikepdf")
        sys.exit(1)

    # Check if the PDF is actually encrypted first
    try:
        with pikepdf.open(path) as pdf:
            # Opened without password — not encrypted
            return path, False
    except pikepdf.PasswordError:
        pass

    password = getpass.getpass("Mandiri PDF password: ")
    try:
        with pikepdf.open(path, password=password) as pdf:
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            pdf.save(tmp.name)
            return tmp.name, True
    except pikepdf.PasswordError:
        print("Error: incorrect password.")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Generate monthly financial diary from bank statements.")
    parser.add_argument("--bca", help="Path to BCA e-statement PDF")
    parser.add_argument("--mandiri", help="Path to Mandiri e-statement (PDF or XLSX)")
    parser.add_argument("--output", default="financial_diary.xlsx", help="Output Excel file path")
    args = parser.parse_args()

    if not args.bca and not args.mandiri:
        print("Error: provide at least one statement (--bca and/or --mandiri).")
        sys.exit(1)

    transactions = []

    if args.bca:
        print(f"Parsing BCA PDF: {args.bca}")
        txns = parse_bca_pdf(args.bca)
        print(f"  → {len(txns)} transactions found")
        transactions.extend(txns)

    if args.mandiri:
        path = args.mandiri
        print(f"Parsing Mandiri Excel: {path}")
        txns = parse_mandiri_excel(path)
        print(f"  → {len(txns)} transactions found")
        transactions.extend(txns)

    if not transactions:
        print("No transactions were parsed. Check your input files.")
        sys.exit(1)

    # Sort by date
    transactions.sort(key=lambda t: t["date"])

    print(f"\nTotal transactions: {len(transactions)}")
    print(f"Writing diary to: {args.output}")
    write_diary(transactions, args.output)
    print("Done.")


if __name__ == "__main__":
    main()