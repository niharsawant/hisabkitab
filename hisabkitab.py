#!/usr/bin/env python3
"""Convert an ICICI family account Excel statement to the HisabKitab CSV format.

Usage: python hisabkitab.py [--dev]
"""

from __future__ import annotations

import csv
import argparse
import importlib
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
HEADERS = ("Date", "Amount", "Category", "Payment Mode", "Type", "Note")
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


def convert(input_path: Path, output_path: Path) -> int:
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
            record = (transaction_date, amount, categorize_note(note), "Family", transaction_type, note)
            transactions.append((date_sort_key(transaction_date), record))
            count += 1
        transactions.sort(key=lambda item: item[0], reverse=True)
        writer.writerows(record for _, record in transactions)
    return count


def main() -> int:
    try:
        parser = argparse.ArgumentParser(description=__doc__)
        parser.add_argument("-d", "--dev", action="store_true", help="use ./input/family.xls without prompting")
        args = parser.parse_args()

        if args.dev:
            input_path = Path.cwd() / "input" / "family.xls"
        else:
            input_value = input("Provide Excel sheet of the family bank account: ").strip()
            if not input_value:
                raise ValueError("No Excel file path was provided")
            input_path = local_path(input_value)
        output_path = Path.cwd() / "output" / f"family-{date.today():%Y-%m-%d}.csv"
        count = convert(input_path, output_path)
    except (OSError, ValueError, KeyError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote {count} transactions to {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
