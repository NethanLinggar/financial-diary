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
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def _decrypt_pdf(path):
    """Prompt for password and return a decrypted copy as a temp file path."""
    try:
        import pikepdf
    except ImportError:
        print("Error: pikepdf is required for password-protected PDFs. Run: pip install pikepdf")
        sys.exit(1)

    try:
        with pikepdf.open(path) as pdf:
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


def _decrypt_excel(path):
    """
    Check if an Excel file is password-protected.
    If so, prompt for the password and return a decrypted file-like object.
    Returns (stream_or_path, was_encrypted).
    """
    try:
        import msoffcrypto
    except ImportError:
        print("Error: msoffcrypto-tool is required for password-protected Excel files. Run: pip install msoffcrypto-tool")
        sys.exit(1)

    with open(path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)

        if not office_file.is_encrypted():
            return path, False

        password = getpass.getpass(f"Excel password for '{os.path.basename(path)}': ")
        try:
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return decrypted, True
        except Exception:
            print("Error: incorrect password or failed to decrypt Excel file.")
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
    tmp_files = []  # track temp files to clean up

    if args.bca:
        print(f"Parsing BCA PDF: {args.bca}")
        txns = parse_bca_pdf(args.bca)
        print(f"  → {len(txns)} transactions found")
        transactions.extend(txns)

    if args.mandiri:
        path = args.mandiri
        ext = os.path.splitext(path)[1].lower()

        if ext in (".xlsx", ".xls", ".xlsm"):
            print(f"Parsing Mandiri Excel: {path}")
            source, was_encrypted = _decrypt_excel(path)
            if was_encrypted:
                print("  (decrypted successfully)")
            txns = parse_mandiri_excel(source)
        else:
            # Assume PDF
            print(f"Parsing Mandiri PDF: {path}")
            source, was_encrypted = _decrypt_pdf(path)
            if was_encrypted:
                tmp_files.append(source)
            txns = parse_mandiri_excel(source)  # swap for parse_mandiri_pdf if you have one

        print(f"  → {len(txns)} transactions found")
        transactions.extend(txns)

    if not transactions:
        print("No transactions were parsed. Check your input files.")
        sys.exit(1)

    transactions.sort(key=lambda t: t["date"])

    print(f"\nTotal transactions: {len(transactions)}")
    print(f"Writing diary to: {args.output}")
    write_diary(transactions, args.output)
    print("Done.")

    # Clean up any decrypted temp PDF files
    for tmp in tmp_files:
        try:
            os.unlink(tmp)
        except OSError:
            pass


if __name__ == "__main__":
    main()