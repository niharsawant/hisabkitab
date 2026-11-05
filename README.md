# HisabKitab

Convert ICICI family and personal bank statements, plus a personal credit card statement, into date-sorted CSV files.

## Setup

The script supports `.xls` and `.xlsx` workbooks. Install its dependencies and create your local configuration:

```sh
python -m pip install python-calamine tomli google-api-python-client google-auth
cp config.example.toml config.toml
```

`tomli` is only needed with Python versions earlier than 3.11. Edit `config.toml` with your category rules and Google Sheets URLs. The local file is ignored by Git; `config.example.toml` contains fictional sample rules and placeholder URLs.

## Run

```sh
python hisabkitab.py
```

The script asks for each file path or `file://` URL in this order. You may leave a prompt blank; missing files are warned about and skipped. At least one valid file is required.

1. Family bank account
2. Personal bank account
3. Personal credit card

The script writes three CSVs to `output/`, named `family-YYYY-MM-DD.csv`, `personal-YYYY-MM-DD.csv`, and `credit-personal-YYYY-MM-DD.csv` using the run date. When run without `--dev`, it also syncs those transactions to the two Google Sheets configured in `config.toml`.

## Development mode

Place any statements you want to process in `input/` with these filenames:

| Statement | Filename |
| --- | --- |
| Family bank account | `family.xls` |
| Personal bank account | `personal.xls` |
| Personal credit card | `credit-personal.xls` |

Then run:

```sh
python hisabkitab.py --dev
```

The `-d` option is an alias for `--dev`. Missing files are warned about and skipped; at least one of the three files must be present.

`--dev` uses the local statements only and skips all Google Sheets writes.

## Google Sheets connection

The script uses a Google Cloud service account and a downloaded JSON key. This provides persistent, non-interactive access: the script obtains short-lived access tokens from the key each time it runs, without asking you to sign in. The service account can access only the spreadsheets shared with its email.

### 1. Create or select a Google Cloud project

Open the [Google Cloud Console](https://console.cloud.google.com/). Select an existing project or create a project for HisabKitab. In that project, open **APIs & Services → Library**, search for **Google Sheets API**, and enable it. The Sheets API must be enabled for the same project that owns the service account.

### 2. Create a service account and download its JSON key

In the Cloud Console, go to **IAM & Admin → Service Accounts** and create a service account. A descriptive name such as `hisabkitab-sheets` is helpful. The service account does not need a project-wide Editor role to update spreadsheets shared with it; the spreadsheet sharing permission in the next step grants the needed access.

Open the new service account, go to **Keys → Add key → Create new key**, choose **JSON**, and create the key. Google downloads the JSON file once. Keep it private: anyone with this file can act as the service account.

By default, place the downloaded file in the project folder with this exact name:

```text
service-account.json
```

This file is ignored by Git. If you store it elsewhere or use a different filename, set `service_account_file` in `config.toml` to its path. Relative paths are resolved from the project directory; an absolute path can also be used. The script expects the original JSON key downloaded from Google Cloud, not credentials copied into TOML.

### 3. Share each spreadsheet with the service account

Open the JSON key in a local editor and copy the value of `client_email` (it ends in `iam.gserviceaccount.com`). In Google Sheets, share **both** target spreadsheets with that email address and grant it **Editor** access. This is spreadsheet sharing; assigning a Cloud project role alone does not grant access to the spreadsheet files.

The service account does not need to be added as a user in the Google Cloud project that owns the spreadsheets. If your Google Workspace administrator restricts sharing outside your organization, they may need to allow the service account email first.

### 4. Configure the two spreadsheet URLs

If you have not created local configuration yet, copy the example:

```sh
cp config.example.toml config.toml
```

Edit the `[google_sheets]` section in `config.toml`. Put the personal spreadsheet URL in `personal_and_credit_url` and the family spreadsheet URL in `family_url`:

```toml
[google_sheets]
personal_and_credit_url = "https://docs.google.com/spreadsheets/d/YOUR_PERSONAL_SHEET_ID/edit?gid=0"
family_url = "https://docs.google.com/spreadsheets/d/YOUR_FAMILY_SHEET_ID/edit?gid=0"
duplicate_window = 10
service_account_file = "service-account.json"
```

Replace the example IDs with the IDs from your actual spreadsheet URLs. The value after `/d/` identifies the spreadsheet. `gid` identifies the worksheet tab; the example uses the first tab (`gid=0`). If you want a different tab, open it in Sheets and use the `gid` shown in its URL. Keep `config.toml` and the JSON key private; neither should be committed.

The script looks for a header row in the first 10 rows. It recognizes aliases such as `Transaction Date`, `Notes`, and `Details`; any missing HisabKitab columns (`Date`, `Amount`, `Category`, `Payment Mode`, `Type`, `Note`, and `Transaction Reference`) are added after the existing columns when syncing. Existing custom columns and their order are preserved. An empty worksheet can be initialized with the standard columns automatically. A non-empty worksheet needs at least one recognizable transaction header in its first 10 rows so the script can identify where to write.

### 5. Run a sync

Run normally to enter statement paths and generate CSVs:

```sh
python hisabkitab.py
```

Normal mode writes CSVs for the valid statements you provide, then syncs those transactions to the configured sheet or sheets. Dates in generated CSVs and new Google Sheets rows use `YYYY-MM-DD`. When either personal statement is processed, any personal bank and personal credit card CSVs already present in `output/` for that run date are combined, sorted newest first, and prepended to the personal spreadsheet. Family transactions are prepended to the family spreadsheet. Existing header labels and formatting are preserved; only missing header cells are added at the end. New transaction rows inherit row formatting from the first existing data row, and the script writes values without replacing cell formatting.

Before inserting, the script fetches the latest `duplicate_window` records below the Google Sheet header (10 by default) and checks whether each appears anywhere in that sheet's generated CSV batch. Personal bank and credit card CSVs are merged for the personal sheet; the family CSV is checked against the family sheet. A matching non-empty transaction reference identifies a duplicate; when a reference is unavailable, it compares Date, Note, and Amount. Matching local rows are skipped. If all fetched sheet records appear in the corresponding CSV batch, no sheet is updated and the script raises an error suggesting that you increase `duplicate_window` in `config.toml` to check more records. It also removes duplicates within the local personal bank and credit card batch.

Development mode does not write to Google Sheets:

```sh
python hisabkitab.py --dev
```

If the service-account JSON is missing or malformed, the script warns and skips the sync while keeping the generated CSVs. Other sync errors, such as a spreadsheet permission error, are reported as errors; first check that the spreadsheet has been shared with the JSON key's `client_email` as an Editor and that the API is enabled. See Google's guides for [creating credentials](https://developers.google.com/workspace/guides/create-credentials), [service-account keys](https://cloud.google.com/iam/docs/keys-create-delete), and the [Sheets API for Python](https://developers.google.com/workspace/sheets/api/quickstart/python).

## CSV and category rules

All files have these columns: `Date`, `Amount`, `Category`, `Payment Mode`, `Type`, `Note`, and `Transaction Reference`. Dates are written as `YYYY-MM-DD`, and transactions are sorted newest first. Bank statement rows use headers on row 13 and data from row 14; credit card rows use headers on row 16 and data from row 18. Credit card `Dr.` amounts are negative Expenses; `Cr.` amounts are positive Income. The credit card `Reference Number` is written to `Transaction Reference`; this field is blank for bank statements.

The script assigns categories by matching each rule's keywords against the transaction note, without regard to case. Rules in `config.toml` are checked top to bottom; the first match wins. Unmatched transactions receive `Others`.

Add one `[[rules]]` section per category in `config.toml`, with a `keys` array and `category` name. For example:

```toml
[[rules]]
keys = ["sample cafe", "demo restaurant"]
category = "Food"
```
