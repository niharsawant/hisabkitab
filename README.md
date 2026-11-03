# hisabkitab
A Python-based personal finance tool that parses and consolidates bank statements from multiple accounts into a unified expense tracker, ready to be exported to Google Sheets. One ledger to make sense of all your money.

## Convert an ICICI family account statement

Install the Excel reader dependency, then run the script. It prompts for the statement path or a `file://` URL. The reader supports both legacy `.xls` and newer `.xlsx` files. The CSV is written to `./output/family-<YYYY-MM-DD>.csv`, using the date the script runs. That folder is ignored by Git.

```sh
python -m pip install python-calamine
python hisabkitab.py
```

When prompted, enter a local path such as `./ICICI statement.xlsx` or a `file://` URL such as `file:///Users/me/ICICI%20statement.xlsx`.

The converter reads the active worksheet, uses the headers in row 13 and transaction rows from row 14 onward, and writes `Date, Amount, Category, Payment Mode, Type, Note`. Withdrawals become negative expenses; deposits become positive income. Category is blank and Payment Mode is `Family`.
