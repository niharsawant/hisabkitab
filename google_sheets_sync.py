"""Sync generated transaction CSVs to the configured Google Sheets."""

from __future__ import annotations

import csv
import re
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Optional
from urllib.parse import parse_qs, urlparse


SCOPES = ("https://www.googleapis.com/auth/spreadsheets",)
EXPECTED_HEADERS = (
    "Date",
    "Amount",
    "Category",
    "Payment Mode",
    "Type",
    "Note",
    "Transaction Reference",
)
HEADER_ALIASES = {
    "date": ("date", "transaction date"),
    "amount": ("amount",),
    "category": ("category",),
    "payment mode": ("payment mode", "paymentmode"),
    "type": ("type",),
    "note": ("note", "notes", "details"),
    "transaction reference": ("transaction reference", "reference number", "reference no"),
}
CANONICAL_HEADERS = dict(zip(HEADER_ALIASES, EXPECTED_HEADERS))
MAX_COLUMNS_RANGE = "AZ"


class InvalidServiceAccountConfig(RuntimeError):
    """Raised when Sheets sync is requested before valid credentials are configured."""


class DuplicateRecordsFound(RuntimeError):
    """Raised when the complete remote comparison window is already in the local CSV batch."""


def normalize_header(value: object) -> str:
    return " ".join(str(value or "").strip().casefold().replace("_", " ").split())


def parse_sheet_url(url: str) -> tuple[str, int]:
    parsed = urlparse(url)
    match = re.search(r"/spreadsheets/d/([^/]+)", parsed.path)
    if not match:
        raise ValueError(f"Not a Google Sheets URL: {url}")
    params = parse_qs(parsed.query)
    if not params.get("gid") and parsed.fragment:
        params = parse_qs(parsed.fragment)
    gid_value = params.get("gid", ["0"])[0]
    try:
        gid = int(gid_value)
    except ValueError as exc:
        raise ValueError(f"Invalid worksheet gid in URL: {url}") from exc
    return match.group(1), gid


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        required = {"Date", "Amount", "Note", "Transaction Reference"}
        missing = required.difference(reader.fieldnames or ())
        if missing:
            raise ValueError(f"{path} is missing required CSV columns: {', '.join(sorted(missing))}")
        return [dict(row) for row in reader]


def parse_date(value: object) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        # Google Sheets date serials use 1899-12-30 as day zero.
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y", "%d/%m/%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        pass
    return None


def fallback_key(row: dict[str, object]) -> Optional[tuple[str, str, str]]:
    transaction_date = parse_date(row.get("Date"))
    note = " ".join(str(row.get("Note") or "").casefold().split())
    if transaction_date is None or not note:
        return None
    raw_amount = str(row.get("Amount") or "").replace(",", "").strip()
    try:
        amount = format(Decimal(raw_amount).normalize(), "f")
    except InvalidOperation:
        amount = raw_amount.casefold()
    return transaction_date.isoformat(), note, amount


class DuplicateIndex:
    """Fast duplicate lookup using a transaction reference or date/note/amount."""

    def __init__(self) -> None:
        self.references: set[str] = set()
        self.fallbacks_with_reference: set[tuple[str, str, str]] = set()
        self.fallbacks_without_reference: set[tuple[str, str, str]] = set()

    @staticmethod
    def normalize_reference(value: object) -> str:
        reference = str(value or "").strip().casefold()
        if re.fullmatch(r"\d+\.0+", reference):
            reference = reference.split(".", 1)[0]
        return reference

    def contains(self, row: dict[str, object]) -> bool:
        reference = self.normalize_reference(row.get("Transaction Reference"))
        key = fallback_key(row)
        if reference:
            return reference in self.references or (key is not None and key in self.fallbacks_without_reference)
        return key is not None and (key in self.fallbacks_with_reference or key in self.fallbacks_without_reference)

    def add(self, row: dict[str, object]) -> None:
        reference = self.normalize_reference(row.get("Transaction Reference"))
        if reference:
            self.references.add(reference)
        key = fallback_key(row)
        if key is not None:
            if reference:
                self.fallbacks_with_reference.add(key)
            else:
                self.fallbacks_without_reference.add(key)


def deduplicate_local(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    """Sort and deduplicate a local batch before comparing it with a sheet."""
    ordered_rows = sorted(rows, key=lambda row: parse_date(row.get("Date")) or date.min, reverse=True)
    index = DuplicateIndex()
    unique = []
    for row in ordered_rows:
        if index.contains(row):
            continue
        index.add(row)
        unique.append(row)
    return unique


def target_field_indexes(headers: list[object]) -> tuple[dict[str, int], list[tuple[int, str]]]:
    indexes: dict[str, int] = {}
    new_headers: list[tuple[int, str]] = []
    normalized = [normalize_header(value) for value in headers]
    for field, aliases in HEADER_ALIASES.items():
        for index, header in enumerate(normalized):
            if header in aliases:
                indexes[field] = index
                break

    # Add any missing HisabKitab fields after the existing header values so
    # custom columns in a user's sheet are preserved.
    next_column = max(
        (index for index, value in enumerate(headers) if str(value or "").strip()),
        default=-1,
    ) + 1
    for field, canonical_header in CANONICAL_HEADERS.items():
        if field not in indexes:
            indexes[field] = next_column
            new_headers.append((next_column, canonical_header))
            next_column += 1
    return indexes, new_headers


def sheet_row_to_record(row: list[object], indexes: dict[str, int]) -> dict[str, object]:
    record: dict[str, object] = {}
    for field in HEADER_ALIASES:
        column = indexes.get(field)
        record[CANONICAL_HEADERS[field]] = row[column] if column is not None and column < len(row) else ""
    return record


def a1_sheet_title(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def locate_sheet(service: Any, spreadsheet_id: str, gid: int) -> dict[str, Any]:
    metadata = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title,gridProperties(columnCount)))",
    ).execute()
    for sheet in metadata.get("sheets", []):
        properties = sheet.get("properties", {})
        if int(properties.get("sheetId", -1)) == gid:
            return properties
    raise ValueError(f"Worksheet gid={gid} was not found in spreadsheet {spreadsheet_id}")


def get_header_and_recent_rows(
    service: Any,
    spreadsheet_id: str,
    title: str,
    recent_count: int = 10,
) -> tuple[int, list[object], list[list[object]]]:
    prefix = a1_sheet_title(title)
    first_rows = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{prefix}!A1:{MAX_COLUMNS_RANGE}10",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute().get("values", [])
    header_index = None
    headers: list[object] = []
    for index, values in enumerate(first_rows):
        normalized = [normalize_header(value) for value in values]
        if any(
            value in aliases
            for value in normalized
            for aliases in HEADER_ALIASES.values()
        ):
            header_index = index
            headers = list(values)
            break
    if header_index is None:
        if not first_rows or not any(str(value or "").strip() for value in first_rows[0]):
            # A completely empty worksheet can start with the standard header row.
            header_index = 0
            headers = []
        else:
            raise ValueError(
                f"Could not identify a header row in worksheet {title!r}. "
                "Add at least one recognized transaction header or clear the first row."
            )

    start_row = header_index + 2  # A1 notation; immediately after the header.
    end_row = start_row + max(1, recent_count) - 1
    recent = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{prefix}!A{start_row}:{MAX_COLUMNS_RANGE}{end_row}",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute().get("values", [])
    return header_index, headers, recent


def make_cell(value: object, field: Optional[str] = None) -> dict[str, object]:
    if value is None or value == "":
        return {}
    if field == "date":
        transaction_date = parse_date(value)
        if transaction_date is None:
            raise ValueError(f"Invalid transaction date for Google Sheets: {value!r}")
        return {"userEnteredValue": {"stringValue": transaction_date.isoformat()}}
    if field == "amount":
        try:
            return {"userEnteredValue": {"numberValue": float(Decimal(str(value).replace(",", "").strip()))}}
        except InvalidOperation as exc:
            raise ValueError(f"Invalid amount for Google Sheets: {value!r}") from exc
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return {"userEnteredValue": {"numberValue": float(value)}}
    return {"userEnteredValue": {"stringValue": str(value)}}


def prepend_unique_rows(
    service: Any,
    spreadsheet_id: str,
    sheet_properties: dict[str, Any],
    header_index: int,
    headers: list[object],
    recent_rows: list[list[object]],
    local_rows: list[dict[str, str]],
    duplicate_window: int = 10,
) -> int:
    indexes, new_headers = target_field_indexes(headers)
    duplicates = DuplicateIndex()
    for row in recent_rows[:duplicate_window]:
        if any(value not in (None, "") for value in row):
            duplicates.add(sheet_row_to_record(row, indexes))

    ordered_local = deduplicate_local(local_rows)
    rows_to_insert = []
    for row in ordered_local:
        if duplicates.contains(row):
            continue
        rows_to_insert.append(row)
    sheet_id = int(sheet_properties["sheetId"])
    data_start = header_index + 1
    column_count = max(len(headers), max(indexes.values(), default=-1) + 1)
    grid_columns = int(sheet_properties.get("gridProperties", {}).get("columnCount", 0))
    requests: list[dict[str, Any]] = []
    if column_count > grid_columns:
        requests.append({"appendDimension": {"sheetId": sheet_id, "dimension": "COLUMNS", "length": column_count - grid_columns}})
    if rows_to_insert:
        requests.append({
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": data_start,
                    "endIndex": data_start + len(rows_to_insert),
                },
                # New rows are below the header, so inherit formatting and
                # dimension properties from the first existing data row.
                "inheritFromBefore": False,
            }
        })
    # Write only missing header cells; existing labels and formatting stay intact.
    for new_header_column, new_header in new_headers:
        requests.append({
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": header_index,
                    "endRowIndex": header_index + 1,
                    "startColumnIndex": new_header_column,
                    "endColumnIndex": new_header_column + 1,
                },
                "rows": [{"values": [make_cell(new_header)]}],
                # Values-only updates do not replace cell formatting.
                "fields": "userEnteredValue",
            }
        })

    if rows_to_insert:
        values = []
        for record in rows_to_insert:
            canonical_values = {
                "date": record.get("Date", ""),
                "amount": record.get("Amount", ""),
                "category": record.get("Category", ""),
                "payment mode": record.get("Payment Mode", ""),
                "type": record.get("Type", ""),
                "note": record.get("Note", ""),
                "transaction reference": record.get("Transaction Reference", ""),
            }
            output_row: list[object] = [""] * column_count
            field_by_column: dict[int, str] = {}
            for field, column in indexes.items():
                if column < column_count:
                    output_row[column] = canonical_values[field]
                    field_by_column[column] = field
            values.append({"values": [make_cell(value, field_by_column.get(column)) for column, value in enumerate(output_row)]})
        requests.append({
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": data_start,
                    "endRowIndex": data_start + len(values),
                    "startColumnIndex": 0,
                    "endColumnIndex": column_count,
                },
                "rows": values,
                # Preserve inherited cell formatting, number formats, and styles.
                "fields": "userEnteredValue",
            }
        })
    if requests:
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
    return len(rows_to_insert)


def make_service(service_account_path: Path) -> Any:
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Missing Google Sheets dependencies. Install them with "
            "'python -m pip install google-api-python-client google-auth'."
        ) from exc
    if not service_account_path.is_file():
        raise InvalidServiceAccountConfig(f"Service account JSON key not found: {service_account_path}")
    try:
        credentials = Credentials.from_service_account_file(str(service_account_path), scopes=SCOPES)
    except Exception as exc:
        raise InvalidServiceAccountConfig(f"Could not load service account JSON key: {service_account_path}") from exc
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def get_spreadsheet_context(
    service: Any,
    url: str,
    duplicate_window: int,
) -> tuple[str, dict[str, Any], int, list[object], list[list[object]]]:
    spreadsheet_id, gid = parse_sheet_url(url)
    properties = locate_sheet(service, spreadsheet_id, gid)
    header_index, headers, recent_rows = get_header_and_recent_rows(
        service, spreadsheet_id, properties["title"], duplicate_window
    )
    return spreadsheet_id, properties, header_index, headers, recent_rows


def all_sheet_records_are_in_local(
    headers: list[object],
    recent_rows: list[list[object]],
    local_rows: list[dict[str, str]],
    duplicate_window: int,
) -> bool:
    """Return true when every non-empty row in a full remote window occurs locally."""
    indexes, _ = target_field_indexes(headers)
    remote_records = [
        sheet_row_to_record(row, indexes)
        for row in recent_rows[:duplicate_window]
        if any(value not in (None, "") for value in row)
    ]
    if len(remote_records) != duplicate_window:
        return False

    local_index = DuplicateIndex()
    for row in local_rows:
        local_index.add(row)
    return all(local_index.contains(record) for record in remote_records)


def sync_exports(
    config: dict[str, Any],
    family_csv: Optional[Path],
    personal_csv: Optional[Path],
    credit_card_csv: Optional[Path],
    config_dir: Path,
) -> dict[str, int]:
    sync_personal = personal_csv is not None or credit_card_csv is not None
    sync_family = family_csv is not None
    if not sync_personal and not sync_family:
        return {}

    sheet_config = config.get("google_sheets")
    if not isinstance(sheet_config, dict):
        raise ValueError("Add a [google_sheets] section to config.toml before running without --dev")
    personal_url = str(sheet_config.get("personal_and_credit_url", "")).strip()
    family_url = str(sheet_config.get("family_url", "")).strip()
    if sync_personal and not personal_url:
        raise ValueError("Set personal_and_credit_url in config.toml")
    if sync_family and not family_url:
        raise ValueError("Set family_url in config.toml")

    key_value = str(sheet_config.get("service_account_file", "service-account.json")).strip()
    key_path = Path(key_value).expanduser()
    if not key_path.is_absolute():
        key_path = config_dir / key_path
    service = make_service(key_path)

    personal_rows = []
    if personal_csv is not None:
        personal_rows.extend(read_csv_rows(personal_csv))
    if credit_card_csv is not None:
        personal_rows.extend(read_csv_rows(credit_card_csv))
    family_rows = read_csv_rows(family_csv) if family_csv is not None else []
    window = int(sheet_config.get("duplicate_window", 10))
    if window < 1:
        raise ValueError("google_sheets.duplicate_window must be a positive integer")

    try:
        plans: dict[str, tuple[str, dict[str, Any], int, list[object], list[list[object]], list[dict[str, str]]]] = {}
        if sync_personal:
            context = get_spreadsheet_context(service, personal_url, window)
            plans["personal and credit card"] = (*context, personal_rows)
        if sync_family:
            context = get_spreadsheet_context(service, family_url, window)
            plans["family"] = (*context, family_rows)

        # Preflight every target before any writes so a duplicate error cannot
        # leave only one of the configured sheets updated.
        for sheet_name, (_, _, _, headers, recent_rows, rows) in plans.items():
            if all_sheet_records_are_in_local(headers, recent_rows, rows, window):
                raise DuplicateRecordsFound(
                    f"Duplicates found: all {window} recent records in the {sheet_name} sheet "
                    "are present in the generated CSVs. No Google Sheets updates were made. "
                    "Increase duplicate_window in config.toml to check more records."
                )

        added: dict[str, int] = {}
        for sheet_name, (spreadsheet_id, properties, header_index, headers, recent_rows, rows) in plans.items():
            added[sheet_name] = prepend_unique_rows(
                service,
                spreadsheet_id,
                properties,
                header_index,
                headers,
                recent_rows,
                rows,
                window,
            )
    except Exception as exc:
        raise RuntimeError(f"Google Sheets synchronization failed: {exc}") from exc
    return added
