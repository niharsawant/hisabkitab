#!/usr/bin/env python3
"""Convert an ICICI family account Excel statement to the HisabKitab CSV format.

Usage: python hisabkitab.py [--dev]
"""

from __future__ import annotations

import csv
import argparse
import importlib
from io import BytesIO
import sys
from datetime import date, datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

try:
    CalamineWorkbook = importlib.import_module("python_calamine").CalamineWorkbook
except ImportError as exc:  # pragma: no cover - exercised by command-line use
    raise SystemExit("Missing dependency: install python-calamine with 'python -m pip install python-calamine'") from exc


FIRST_DATA_ROW = 14
FIRST_COLUMN = 2  # B
LAST_COLUMN = 9  # I
CARD_HEADER_ROW = 16
CARD_FIRST_DATA_ROW = 18
CARD_FIRST_COLUMN = 2  # B
CARD_LAST_COLUMN = 18  # R
HEADERS = ("Date", "Amount", "Category", "Payment Mode", "Type", "Note", "Transaction Reference")
__version__ = "v0.1.0"
CATEGORY_RULES = (
    (("jio", "tata power", "mahanagar", "tata play"), "Bills"),
    (("zomato", "heisetasse", "third wave", "starbucks", "bluetokai", "brijwasi", "blue tokai", "food"), "Food"),
    (("blinkit", "grofers", "licious", "naturesbasket", "natures bas", "grocery", "grn organic", "mahboob450", "rsp blink"), "Grocery"),
    (("rentomojo", "amazon", "rajkumar"), "Home"),
    (("salvipradnya12", "dhotrehanuman5"), "Househelp"),
    (("nanny", "vrushalichavan", "rushali"), "Nanny"),
    (("netflix",), "Entertainment"),
    (("family",), "Contribution"),
    (("rent",), "Rent"),
    (("dialpad",), "Salary"),
    (("uber", "ola", "auto", "rickshaw"), "Local Travel"),
    (("mother chemist", "medicine"), "Pharmacy"),
    (("firstcry",), "Childcare")
)


def local_path(value: str) -> Path:
    """Accept a filesystem path or a file:// URL."""
    parsed = urlparse(value)
    if parsed.scheme and parsed.scheme != "file":
        raise ValueError("input must be a local path or a file:// URL")
    if parsed.scheme == "file":
        path = unquote(parsed.path)
        if parsed.netloc and parsed.netloc != "localhost":
            path = f"//{parsed.netloc}{path}"
        return Path(path).expanduser().resolve()
    return Path(value).expanduser().resolve()


def normalize_header(value: object) -> str:
    return " ".join(str(value or "").strip().lower().replace("\n", " ").split())


def find_column(headers: list[str], names: tuple[str, ...]) -> int:
    """Return an absolute worksheet column for one of the expected headers."""
    normalized = [normalize_header(value) for value in headers]
    for index, header in enumerate(normalized):
        if any(name in header for name in names):
            return FIRST_COLUMN + index
    raise ValueError(f"Could not find a column matching: {', '.join(names)}")


def date_sort_key(value: object) -> tuple[int, date]:
    """Parse common statement date values; unparseable dates sort last."""
    if isinstance(value, datetime):
        return (1, value.date())
    if isinstance(value, date):
        return (1, value)
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return (1, datetime.strptime(text, fmt).date())
        except ValueError:
            pass
    return (0, date.min)


def categorize_note(note: object) -> str:
    text = str(note or "").casefold()
    for keys, category in CATEGORY_RULES:
        if any(key in text for key in keys):
            return category
    return "Others"


def convert(input_path: Path, output_path: Path, payment_mode: str) -> int:
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    workbook = CalamineWorkbook.from_path(input_path)
    sheet = workbook.get_sheet_by_index(0).to_python(skip_empty_area=False)
    headers = [sheet[12][column - 1] if len(sheet) > 12 and len(sheet[12]) >= column else None
               for column in range(FIRST_COLUMN, LAST_COLUMN + 1)]
    date_col = find_column(headers, ("transaction date", "txn date", "date")) - 1
    withdrawal_col = find_column(headers, ("withdrawal", "debit")) - 1
    deposit_col = find_column(headers, ("deposit", "credit")) - 1
    remark_col = find_column(headers, ("transaction remark", "transaction description", "narration", "particulars", "description")) - 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    transactions = []
    count = 0
    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(HEADERS)
        for values in sheet[FIRST_DATA_ROW - 1:]:
            row = list(values) + [None] * max(0, LAST_COLUMN - len(values))
            transaction_date = row[date_col] if len(row) > date_col else None
            withdrawal = row[withdrawal_col] if len(row) > withdrawal_col else None
            deposit = row[deposit_col] if len(row) > deposit_col else None
            remark = row[remark_col] if len(row) > remark_col else None

            # Ignore trailing blank rows and rows that contain no transaction data.
            if all(value is None or str(value).strip() == "" for value in (transaction_date, withdrawal, deposit, remark)):
                continue
            if withdrawal is not None and str(withdrawal).strip() != "" and float(withdrawal) == 0:
                if deposit is None or str(deposit).strip() == "":
                    continue
                amount = float(deposit)
                transaction_type = "Income"
            elif withdrawal is not None and str(withdrawal).strip() != "":
                amount = -abs(float(withdrawal))
                transaction_type = "Expense"
            elif deposit is not None and str(deposit).strip() != "":
                amount = float(deposit)
                transaction_type = "Income"
            else:
                # Skip non-transaction rows such as repeated headings or balances.
                continue

            if hasattr(transaction_date, "strftime"):
                transaction_date = transaction_date.strftime("%Y-%m-%d")
            note = remark or ""
            record = (transaction_date, amount, categorize_note(note), payment_mode, transaction_type, note, "")
            transactions.append((date_sort_key(transaction_date), record))
            count += 1
        transactions.sort(key=lambda item: item[0], reverse=True)
        writer.writerows(record for _, record in transactions)
    return count


def convert_credit_card(input_path: Path, output_path: Path) -> int:
    """Convert the personal credit card statement layout (headers row 16, data row 18)."""
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    workbook = CalamineWorkbook.from_filelike(BytesIO(input_path.read_bytes()))
    sheet = workbook.get_sheet_by_index(0).to_python(skip_empty_area=False)
    header_index = CARD_HEADER_ROW - 1
    headers = [
        sheet[header_index][column - 1]
        if len(sheet) > header_index and len(sheet[header_index]) >= column
        else None
        for column in range(CARD_FIRST_COLUMN, CARD_LAST_COLUMN + 1)
    ]
    date_col = find_column(headers, ("transaction date", "date")) - 1
    details_col = find_column(headers, ("details", "description", "narration")) - 1
    amount_col = find_column(headers, ("amount",)) - 1
    reference_col = find_column(headers, ("reference number", "reference no")) - 1

    transactions = []
    for values in sheet[CARD_FIRST_DATA_ROW - 1:]:
        row = list(values) + [None] * max(0, CARD_LAST_COLUMN - len(values))
        transaction_date = row[date_col] if len(row) > date_col else None
        details = row[details_col] if len(row) > details_col else None
        raw_amount = row[amount_col] if len(row) > amount_col else None
        reference = row[reference_col] if len(row) > reference_col else None
        if raw_amount is None or str(raw_amount).strip() == "":
            continue

        amount_text = str(raw_amount).strip()
        is_debit = amount_text.casefold().endswith("dr.")
        is_credit = amount_text.casefold().endswith("cr.")
        if is_debit or is_credit:
            amount_text = amount_text[:-3].strip()
        amount_text = amount_text.replace(",", "").replace("INR", "").strip()
        amount_value = float(raw_amount) if isinstance(raw_amount, (int, float)) else float(amount_text)
        if is_debit:
            amount = -abs(amount_value)
            transaction_type = "Expense"
        else:
            amount = abs(amount_value)
            transaction_type = "Income"

        if hasattr(transaction_date, "strftime"):
            transaction_date = transaction_date.strftime("%Y-%m-%d")
        note = details or ""
        record = (transaction_date, amount, categorize_note(note), "Personal Credit Card", transaction_type, note, reference or "")
        transactions.append((date_sort_key(transaction_date), record))

    transactions.sort(key=lambda item: item[0], reverse=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(HEADERS)
        writer.writerows(record for _, record in transactions)
    return len(transactions)


def main() -> int:
    try:
        parser = argparse.ArgumentParser(description=__doc__)
        parser.add_argument("-d", "--dev", action="store_true", help="use the three statement files in ./input without prompting")
        args = parser.parse_args()

        if args.dev:
            input_paths = (
                Path.cwd() / "input" / "family.xls",
                Path.cwd() / "input" / "personal.xls",
                Path.cwd() / "input" / "credit-personal.xls",
            )
        else:
            account_inputs = (
                ("family", "Provide Excel sheet of the family bank account: "),
                ("personal", "Provide Excel sheet of the personal bank account: "),
                ("personal credit card", "Provide Excel sheet of the personal credit card statement: "),
            )
            input_paths = []
            for account, prompt in account_inputs:
                input_value = input(prompt).strip()
                if not input_value:
                    raise ValueError(f"No Excel file path was provided for the {account} account")
                input_paths.append(local_path(input_value))

        output_dir = Path.cwd() / "output"
        run_date = date.today().strftime("%Y-%m-%d")
        output_paths = (
            output_dir / f"family-{run_date}.csv",
            output_dir / f"personal-{run_date}.csv",
            output_dir / f"credit-personal-{run_date}.csv",
        )
        payment_modes = ("Family", "Personal")
        missing_inputs = [str(path) for path in input_paths if not path.is_file()]
        if missing_inputs:
            raise FileNotFoundError(f"Excel file not found: {', '.join(missing_inputs)}")
        for input_path, output_path, payment_mode in zip(input_paths[:2], output_paths[:2], payment_modes):
            count = convert(input_path, output_path, payment_mode)
            print(f"Wrote {count} transactions to {output_path}")
        count = convert_credit_card(input_paths[2], output_paths[2])
        print(f"Wrote {count} transactions to {output_paths[2]}")
    except (OSError, ValueError, KeyError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
