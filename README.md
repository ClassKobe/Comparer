# EBuilder → PRIVV Processor

A desktop (Tkinter) tool for reconciling construction commitment and invoice
data exported from **eBuilder** against **PRIVV** exports, for the Beaver
Stadium Renovation project. It cross-checks vendor amounts, classifies
commitments into PRIVV budget line codes, matches invoices between the two
systems, and exports formatted Excel reports.

## What it does

The app has three tabs:

### 1. Commited (Budget/Commitment Comparison)
- Load one or more eBuilder commitment export CSVs and a single PRIVV
  commitment export CSV.
- Classifies each eBuilder commitment into a PRIVV budget code (e.g. `403`,
  `503.4`, `711`) using a weighted keyword classifier (`CLASSIFICATION_RULES`)
  that looks at vendor name, description, and commitment type.
- Cross-references vendor totals between eBuilder and PRIVV, using fuzzy
  matching (`rapidfuzz`) plus a vendor alias table to reconcile vendors that
  are named differently in each system.
- Displays a side-by-side comparison window (amounts, deltas, match status).


### 2. Invoices
- Load eBuilder invoice export CSV(s) (needs `Commitment #` and
  `Invoice Amount` columns) and a PRIVV invoice export CSV (needs
  `Vendor #` and `Amount` columns).
- Groups PRIVV rows by vendor and matches eBuilder invoice rows to each
  vendor by Company name, using the same alias/fuzzy matching logic as the
  Commited tab, so mismatched or reused vendor numbers don't cause invoices
  to be misassigned.
- Handles a few special cases automatically:
  - Penn State Office of Physical Plant name variants are collapsed into one
    "Penn State Office of Physical Plant" line.
  - Rows that can't be matched by vendor/alias fall back to a fuzzy match
    against Company/Description text before being flagged as unmatched.
- Shows a comparison window (PRIVV total vs. eBuilder total, delta, match
  count, status) with an **Export to Excel** button that produces a
  multi-sheet workbook: a summary sheet plus the raw matched eBuilder and
  PRIVV invoice line items.

### 3. Vendor Aliases
- Maintains a shared "Vendor Name → Alias" table stored in a Google Sheet
  (tab `VendorAliases` inside the `learned_rules` spreadsheet), so multiple
  users can teach the matcher new vendor-name variants without editing code.
- Aliases entered here are combined with the hardcoded `VENDOR_ALIASES` map
  in the script and used immediately by both the Commited and Invoices tabs.
- If `gspread`/`google-auth` aren't installed, or `credentials.json` isn't
  present, this tab still works locally but aliases won't be shared/persisted
  to the sheet.

## Requirements

- Python 3.9+
- Packages:
  ```
  pandas
  rapidfuzz
  openpyxl
  GitPython
  google-auth
  gspread
  ```
  (Tkinter ships with most standard Python installs; `google-auth`/`gspread`
  are optional — the app degrades gracefully to local-only aliases without
  them.)

## Google Sheets setup (optional, for shared Vendor Aliases)

1. Create a Google Cloud service account and download its JSON key.
2. Save the key file as `credentials.json` in the same folder as the script.
3. Share the `learned_rules` Google Sheet with the service account's email
   address (found inside `credentials.json`).
4. The app will create a `VendorAliases` tab in that sheet automatically the
   first time it's needed.

Without this setup, the Vendor Aliases tab and alias matching still function
using only the hardcoded `VENDOR_ALIASES` table in the script.

## Running the app

```bash
python New_Python_Source_File.py
```

This opens the GUI directly — there's no command-line mode.

## Typical workflow

1. **Commited tab:** select your eBuilder commitment CSV(s) and PRIVV
   commitment CSV, then run the comparison. Review any mismatches, add
   vendor aliases as needed (Vendor Aliases tab), and re-run until totals
   reconcile. 
2. **Invoices tab:** select your eBuilder invoice CSV(s) and PRIVV invoice
   CSV, run the comparison, review the results window, and use
   **Export to Excel** to save a formatted workbook for sharing/records.
3. **Vendor Aliases tab:** whenever a vendor is matched incorrectly or not
   matched at all, add a Vendor Name → Alias pair here so future runs
   resolve it correctly.
## Notes
1. Do NOT upload the credentials.json to github or google as it will invalidate the key making it so the key no longer works. Meaning the code will not have access to the list of vendor allias and will make the program potentially throw an error.
2. This code will only work on the PSU project
