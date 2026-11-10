# 📒 HisabKitab

Convert family, personal bank, and credit card statements into date-sorted CSV files, then sync the transactions to Google Sheets.

## ⚙️ Setup

The script reads `.xls` and `.xlsx` statements, plus headerless CSV files for HSBC credit cards. First, install the required packages and make a local copy of the example configuration:

```sh
python3 -m pip install python-calamine google-api-python-client google-auth
cp config.example.toml config.toml
```

If you’re using Python earlier than 3.11, install `tomli` too:

```sh
python3 -m pip install tomli
```

### 📄 Configure the TOML file

Open `config.toml` to set up category rules, statement banks and formats, and Google Sheets options. It’s a local file that may contain private details, so keep it private. `config.example.toml` includes sample rules and placeholder sheet URLs.

Here’s what you can configure:

| Setting | What it does |
| --- | --- |
| `rules.keys` | Text snippets to look for in a transaction note. Matches are case-insensitive. |
| `rules.category` | The category to use when one of the rule’s keys matches. |
| `statements.family_account_bank` | Bank for the family account statement. Currently only `ICICI` is supported. |
| `statements.personal_account_bank` | Bank for the personal account statement. Currently only `ICICI` is supported. |
| `statements.personal_credit_account_bank` | Bank for the personal credit card statement. Supports `ICICI` (Excel) and `HSBC` (CSV). |
| `google_sheets.personal_and_credit_url` | The sheet where personal bank and credit card transactions are combined. |
| `google_sheets.family_url` | The sheet for family bank transactions. |
| `google_sheets.duplicate_window` | How many recent sheet records to check for duplicates. The default is `10`. |
| `google_sheets.service_account_file` | The path to the Google service account JSON key. The default is `service-account.json` in the project folder. |

Add one `[[rules]]` block for each category. For example:

```toml
[[rules]]
keys = ["sample cafe", "demo restaurant"]
category = "Food"

[[rules]]
keys = ["example power company"]
category = "Bills"
```

If none of the keywords match, the transaction is categorized as `Others`.

The three bank settings live in `[statements]`. Family and personal bank accounts currently use `ICICI`. For personal credit cards, use `ICICI` for the existing Excel statement or `HSBC` for the HSBC CSV statement. The HSBC CSV must not have a header row; its first three columns must be date, note, and amount, in that order. Category is inferred from the note using the same rules as the other statements.

### 🔗 Connect to Google Sheets

HisabKitab uses a Google Cloud service account to connect without asking you to sign in on every run. The service account can access only the spreadsheets you share with it.

1. Open the [Google Cloud Console](https://console.cloud.google.com/) and select an existing project or create one.
2. Go to **APIs & Services → Library**, find **Google Sheets API**, and enable it for that project.
3. Open **IAM & Admin → Service Accounts** and create a service account. You could name it `hisabkitab-sheets`. It doesn’t need a project-wide Editor role.
4. Open the service account and choose **Keys → Add key → Create new key → JSON**. Create and download the key.
5. Save the downloaded file in the project folder as `service-account.json`.

🔐 Keep `service-account.json` private. Anyone who has this file can use the service account. Git ignores the file. If you save it under another name or in another location, update `service_account_file` in `config.toml`.

Open the JSON key locally and copy the value of `client_email`. In Google Sheets, share **both** spreadsheets with that email address and give it **Editor** access. A Cloud project role alone doesn’t grant access to your spreadsheets; sharing each sheet does.

### 🧾 Add the sheet URLs to TOML

Add or update this `[google_sheets]` section in `config.toml`:

```toml
[google_sheets]
personal_and_credit_url = "https://docs.google.com/spreadsheets/d/YOUR_PERSONAL_SHEET_ID/edit?gid=0"
family_url = "https://docs.google.com/spreadsheets/d/YOUR_FAMILY_SHEET_ID/edit?gid=0"
duplicate_window = 10
service_account_file = "service-account.json"
```

Replace the sample IDs with the IDs from your spreadsheets. The ID is the part after `/d/` in the URL. `gid` selects a worksheet tab; `gid=0` is commonly the first tab. To use a different tab, open it and copy the `gid` from the URL.

For a sheet that already contains data, its header row should be somewhere in the first 10 rows. HisabKitab recognizes names such as `Date` or `Transaction Date`, and `Note`, `Notes`, or `Details`. Missing HisabKitab columns are added after the existing columns. Existing headers and custom columns are kept in place; an empty sheet can be initialized automatically.

When syncing, the script compares the latest `duplicate_window` records in each sheet with the corresponding generated CSV rows. Personal bank and credit card records are combined for the personal sheet; family records are checked against the family sheet. Matching rows are skipped. If every fetched sheet record is already in its CSV batch, the run stops with an error and suggests increasing `duplicate_window`. This check happens before writing, so neither sheet is updated in that case.

## ▶️ Run the script

Run without `--dev` to enter your statement paths and sync transactions to Google Sheets:

```sh
python3 hisabkitab.py
```

At each prompt, enter a local file path or `file://` URL. The prompts appear in this order:

1. Family bank account
2. Personal bank account
3. Personal credit card statement

You can leave a prompt blank to skip that statement. Missing or invalid files are warned about and skipped; at least one valid file is needed. CSV filenames include the run date, and transaction dates use `YYYY-MM-DD`:

- `family-YYYY-MM-DD.csv`
- `personal-YYYY-MM-DD.csv`
- `credit-personal-YYYY-MM-DD.csv`

Each generated CSV includes a final `Bank` column populated from the matching bank setting in `config.toml`.

In normal mode, CSVs are created in a temporary folder for Google Sheets sync and removed when the script finishes, including when syncing fails. They are not left in `output/`. When you process either personal statement, any counterpart CSV already in `output/` for that date is also included in the personal sheet sync. Family transactions go to the family sheet. If the service account JSON is missing or invalid, you’ll see a warning that syncing was skipped, and the temporary CSVs are still cleaned up. Other sync issues, such as missing sheet access, are reported as errors.

## 🧪 Run in the development environment

During development, keep any statements you want to process in `input/` using these filenames:

| Statement | Filename |
| --- | --- |
| Family bank account | `family.xls` |
| Personal bank account | `personal.xls` |
| Personal credit card with ICICI selected | `credit-personal.xls` |
| Personal credit card with HSBC selected | `credit-personal.csv` |

Then run:

```sh
python3 hisabkitab.py --dev
```

HisabKitab processes whichever of these files are present, as long as at least one is available. It saves the CSVs to `output/` and skips Google Sheets writes in development mode.

## 🧰 Command-line arguments

| Argument | What it does |
| --- | --- |
| `-d`, `--dev` | Reads statement files from `input/` without prompting. Skips Google Sheets writes. |
| `-h`, `--help` | Shows command usage and the available arguments. |

In normal mode, the script prompts for file paths; you don’t pass statement paths as command-line arguments.
