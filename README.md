# hisabkitab
A Python-based personal finance tool that parses and consolidates bank statements from multiple accounts into a unified expense tracker, ready to be exported to Google Sheets. One ledger to make sense of all your money.

## Convert an ICICI family account statement

Install the Excel reader dependency, then run the script. By default, it prompts for the statement path or a `file://` URL. The reader supports both legacy `.xls` and newer `.xlsx` files. The CSV is written to `./output/family-<YYYY-MM-DD>.csv`, using the date the script runs.

```sh
python -m pip install python-calamine
python hisabkitab.py
```

When prompted, enter a local path such as `./ICICI statement.xlsx` or a `file://` URL such as `file:///Users/me/ICICI%20statement.xlsx`.

### Development mode

During development, keep the family account statement in the `input` folder as `family.xls`. Run with `--dev` (or `-d`) to skip the prompt and use that file:

```sh
python hisabkitab.py --dev
# or
python hisabkitab.py -d
```


The converter reads the active worksheet, uses the headers in row 13 and transaction rows from row 14 onward, and writes `Date, Amount, Category, Payment Mode, Type, Note`. Categories are assigned by case-insensitive keyword matches in the transaction note; unmatched notes are categorized as `Others`. Withdrawals become negative expenses; deposits become positive income. Payment Mode is `Family`.
