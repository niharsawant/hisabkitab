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
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import unquote, urlparse

try:
    CalamineWorkbook = importlib.import_module("python_calamine").CalamineWorkbook
except ImportError as exc:  # pragma: no cover - exercised by command-line use
    raise SystemExit("Missing dependency: install python-calamine with 'python -m pip install python-calamine'") from exc

try:
    tomllib = importlib.import_module("tomllib")
except ModuleNotFoundError:  # Python 3.10 and earlier
    try:
        tomllib = importlib.import_module("tomli")
    except ImportError as exc:  # pragma: no cover - exercised by command-line use
        raise SystemExit("Missing dependency: install tomli with 'python -m pip install tomli'") from exc


FIRST_DATA_ROW = 14
FIRST_COLUMN = 2  # B
LAST_COLUMN = 9  # I
CARD_HEADER_ROW = 16
CARD_FIRST_DATA_ROW = 18
CARD_FIRST_COLUMN = 2  # B
CARD_LAST_COLUMN = 18  # R
HEADERS = ("Date", "Amount", "Category", "Payment Mode", "Type", "Note", "Transaction Reference")
__version__ = "v0.1.0"
CATEGORY_RULES_PATH = Path(__file__).with_name("config.toml")


def load_config(path: Path = CATEGORY_RULES_PATH) -> dict[str, object]:
    if not path.is_file():
        example_path = path.with_name("config.example.toml")
        raise SystemExit(f"Category rules file not found: {path}. Copy {example_path.name} to {path.name}.")

    try:
        with path.open("rb") as rules_file:
            config = tomllib.load(rules_file)
    except tomllib.TOMLDecodeError as exc:
        raise SystemExit(f"Invalid TOML in {path}: {exc}") from exc
    return config


def load_category_rules(config: dict[str, object]) -> tuple[tuple[tuple[str, ...], str], ...]:
    rules = config.get("rules")
    if not isinstance(rules, list) or not rules:
        raise SystemExit(f"{CATEGORY_RULES_PATH} must contain one or more [[rules]] entries")

    parsed_rules = []
    for index, rule in enumerate(rules, start=1):
        keys = rule.get("keys") if isinstance(rule, dict) else None
        category = rule.get("category") if isinstance(rule, dict) else None
        if not isinstance(keys, list) or not keys or not all(isinstance(key, str) and key.strip() for key in keys):
            raise SystemExit(f"Rule {index} in {CATEGORY_RULES_PATH} must have a non-empty array of string keys")
        if not isinstance(category, str) or not category.strip():
            raise SystemExit(f"Rule {index} in {CATEGORY_RULES_PATH} must have a non-empty string category")
        parsed_rules.append((tuple(key.strip().casefold() for key in keys), category.strip()))
    return tuple(parsed_rules)


CONFIG = load_config()
CATEGORY_RULES = load_category_rules(CONFIG)


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


def normalize_transaction_date(value: object) -> str:
    """Return a transaction date as YYYY-MM-DD or reject an unrecognized value."""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date().isoformat()
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y", "%d/%m/%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text).date().isoformat()
    except ValueError as exc:
        raise ValueError(f"Unrecognized transaction date: {value!r}; expected a valid date") from exc


def date_sort_key(value: object) -> tuple[int, date]:
    """Sort valid transaction dates newest first; invalid values sort last."""
    try:
        return (1, date.fromisoformat(normalize_transaction_date(value)))
    except ValueError:
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

            transaction_date = normalize_transaction_date(transaction_date)
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

        transaction_date = normalize_transaction_date(transaction_date)
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
            expected_inputs = {
                "family": Path.cwd() / "input" / "family.xls",
                "personal": Path.cwd() / "input" / "personal.xls",
                "credit_personal": Path.cwd() / "input" / "credit-personal.xls",
            }
            input_paths = {}
            for account, input_path in expected_inputs.items():
                if input_path.is_file():
                    input_paths[account] = input_path
                else:
                    print(f"Warning: development input file not found for {account}: {input_path}; skipping it.", file=sys.stderr)
            if not input_paths:
                raise FileNotFoundError("No development input files were found in ./input; nothing to convert.")
        else:
            account_inputs = (
                ("family", "Provide Excel sheet of the family bank account: "),
                ("personal", "Provide Excel sheet of the personal bank account: "),
                ("credit_personal", "Provide Excel sheet of the personal credit card statement: "),
            )
            input_paths = {}
            for account, prompt in account_inputs:
                input_value = input(prompt).strip()
                if not input_value:
                    print(f"Warning: no file provided for {account}; skipping it.", file=sys.stderr)
                    continue
                try:
                    input_path = local_path(input_value)
                except ValueError as exc:
                    print(f"Warning: invalid path for {account} ({exc}); skipping it.", file=sys.stderr)
                    continue
                if not input_path.is_file():
                    print(f"Warning: file not found for {account}: {input_path}; skipping it.", file=sys.stderr)
                    continue
                input_paths[account] = input_path

            if not input_paths:
                raise ValueError("No valid statement files were provided; nothing to convert.")

        output_dir = Path.cwd() / "output"
        run_date = date.today().strftime("%Y-%m-%d")
        output_paths = {
            "family": output_dir / f"family-{run_date}.csv",
            "personal": output_dir / f"personal-{run_date}.csv",
            "credit_personal": output_dir / f"credit-personal-{run_date}.csv",
        }
        payment_modes = {"family": "Family", "personal": "Personal"}
        for account, input_path in input_paths.items():
            output_path = output_paths[account]
            if account == "credit_personal":
                count = convert_credit_card(input_path, output_path)
            else:
                count = convert(input_path, output_path, payment_modes[account])
            print(f"Wrote {count} transactions to {output_path}")
        if not args.dev:
            from google_sheets_sync import InvalidServiceAccountConfig, sync_exports

            sync_personal = "personal" in input_paths or "credit_personal" in input_paths
            personal_csv = output_paths["personal"] if sync_personal and output_paths["personal"].is_file() else None
            credit_personal_csv = (
                output_paths["credit_personal"]
                if sync_personal and output_paths["credit_personal"].is_file()
                else None
            )

            try:
                added = sync_exports(
                    CONFIG,
                    output_paths.get("family") if "family" in input_paths else None,
                    personal_csv,
                    credit_personal_csv,
                    CATEGORY_RULES_PATH.parent,
                )
            except InvalidServiceAccountConfig as exc:
                print(
                    f"Warning: Google Sheets sync was skipped because service account credentials are missing or invalid. {exc} "
                    "The CSV files were generated.",
                    file=sys.stderr,
                )
                return 0
            if "personal and credit card" in added:
                print(f"Added {added['personal and credit card']} rows to the personal and credit card Google Sheet")
            if "family" in added:
                print(f"Added {added['family']} rows to the family Google Sheet")
    except (OSError, ValueError, KeyError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
