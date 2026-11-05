# hisabkitab
A Python-based personal finance tool that parses and consolidates bank statements from multiple accounts into a unified expense tracker, ready to be exported to Google Sheets. One ledger to make sense of all your money.

## Convert an ICICI family account statement

Install the Excel reader dependency, then run the script. By default, it prompts for the family account statement, the personal account statement, and the personal credit card statement, accepting a local path or a `file://` URL for each. The reader supports both legacy `.xls` and newer `.xlsx` files. It writes `./output/family-<YYYY-MM-DD>.csv`, `./output/personal-<YYYY-MM-DD>.csv`, and `./output/credit-personal-<YYYY-MM-DD>.csv`, using the date the script runs.

```sh
python -m pip install python-calamine
python hisabkitab.py
```

When prompted for each statement, enter a local path such as `./ICICI statement.xlsx` or a `file://` URL such as `file:///Users/me/ICICI%20statement.xlsx`.

### Development mode

During development, keep the family account statement in the `input` folder as `family.xls`, the personal account statement as `personal.xls`, and the personal credit card statement as `credit-personal.xls`. Run with `--dev` (or `-d`) to skip the prompts and use those files:

```sh
python hisabkitab.py --dev
# or
python hisabkitab.py -d
```


The bank statement converters read the active worksheet, use headers in row 13 and transaction rows from row 14 onward, and write `Date, Amount, Category, Payment Mode, Type, Note`. The personal credit card converter reads headers from row 16 and transactions from row 18 onward; `Dr.` amounts become negative expenses, and `Cr.` amounts become positive income. Categories use case-insensitive keyword matches in the transaction note; unmatched notes are categorized as `Others`. Payment Mode is `Family`, `Personal`, or `Personal Credit Card` for the corresponding statement.
