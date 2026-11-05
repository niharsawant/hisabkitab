# HisabKitab

Convert ICICI family and personal bank statements, plus a personal credit card statement, into date-sorted CSV files.

## Setup

The script supports `.xls` and `.xlsx` workbooks. Install its dependencies and create your local category configuration:

```sh
python -m pip install python-calamine tomli
cp config.example.toml config.toml
```

`tomli` is only needed with Python versions earlier than 3.11. Edit `config.toml` with your own category keywords. The local file is ignored by Git; `config.example.toml` contains fictional sample rules.

## Run

```sh
python hisabkitab.py
```

The script asks for each statement path or file URL in order. You may leave a prompt blank; missing or invalid files are warned about and skipped. At least one valid file is required.

1. Family bank account
2. Personal bank account
3. Personal credit card

The script writes three CSVs to `output/`, named `family-YYYY-MM-DD.csv`, `personal-YYYY-MM-DD.csv`, and `credit-personal-YYYY-MM-DD.csv` using the run date.

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

## CSV and category rules

All files have these columns: `Date`, `Amount`, `Category`, `Payment Mode`, `Type`, `Note`, and `Transaction Reference`. Transactions are sorted newest first. Bank statement rows use headers on row 13 and data from row 14; credit card rows use headers on row 16 and data from row 18. Credit card `Dr.` amounts are negative Expenses; `Cr.` amounts are positive Income. The credit card `Reference Number` is written to `Transaction Reference`; this field is blank for bank statements.

The script assigns categories by matching each rule's keywords against the transaction note, without regard to case. Rules in `config.toml` are checked top to bottom; the first match wins. Unmatched transactions receive `Others`.

Add one `[[rules]]` section per category in `config.toml`, with a `keys` array and `category` name. For example:

```toml
[[rules]]
keys = ["sample cafe", "demo restaurant"]
category = "Food"
```
