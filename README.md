EBuilder to PRIVV Processor
============================

OVERVIEW

This tool processes exported eBuilder commitment data and maps it to PRIVV-compatible cost codes for import. It combines multiple eBuilder CSV files, classifies each line item using a weighted keyword and fuzzy matching system, and outputs a formatted CSV ready for PRIVV import. It also includes a vendor comparison tool to reconcile dollar amounts between the two systems.


REQUIREMENTS

- Python 3.8 or higher
- The following Python packages:
  - pandas
  - rapidfuzz
  - tkinter (included with most Python installations)

Install dependencies with:
  pip install pandas rapidfuzz


HOW TO RUN

Run the script directly with Python:
  python ebuilder_privv_processor.py

A GUI window will open with the following controls.


USING THE TOOL

Step 1 - Select eBuilder Files
Click "Select eBuilder Files" to choose one or more exported eBuilder CSV files. Multiple files will be combined and deduplicated automatically. Click "Clear" to reset the selection.

Step 2 - Select PRIVV File
Click "Select PRIVV File" to choose your PRIVV reference CSV. This file must contain at minimum a "Vendor" column and an "Amount" column. If it also contains a "Code" column, those codes will be used to supplement the fuzzy matching logic.

Step 3 - Run or Compare

  Run (Export CSV):
  Processes the eBuilder files, classifies each row to a cost code, groups Purchase Order rows by vendor and code, and writes the result to a file called privv_importable.csv in the same directory as the script.

  Compare Vendors:
  Opens a separate window showing a side-by-side dollar comparison of every vendor found in the PRIVV file against the matched totals in eBuilder. Rows are color coded: green for matches, yellow when eBuilder is higher, red when eBuilder is lower, and gray when no data exists on either side. You can filter by vendor name and export the comparison to CSV.


OUTPUT FILE

The output file privv_importable.csv contains the following columns:

- Code            : The assigned PRIVV cost code
- Vendor          : The vendor/company name from eBuilder
- Description     : The commitment description from eBuilder
- PO Number       : Left blank (for manual entry if needed)
- Item Description: The human-readable label for the cost code
- Date Committed  : The commitment date, formatted as M/D/YYYY
- Type            : Either "Original" or "Change Order"
- Amount          : The current commitment dollar amount


COST CODE CLASSIFICATION

Each row is assigned a cost code using a layered approach:

1. Keyword Rules - A list of weighted rules checks the vendor name, description, and commitment type for known keywords. Higher weight rules take priority. For example, any row mentioning "Populous" is always mapped to code 403 (Project Architect Reimbursables).

2. Fuzzy Lookup - If no keyword rule wins with a high enough score, the vendor name is fuzzy-matched against vendors in the PRIVV reference file using token set ratio matching. A minimum score of 40 is required.

3. Default Fallback - If nothing matches, the row is assigned code 901 (Owner Contingency).

Vendor aliases are also applied before classification to normalize common name variations between eBuilder and PRIVV (for example, "Insane Impact LLC" becomes "Insane/Impact").


VENDOR COMPARISON NOTES

- Unmatched eBuilder rows (those with no corresponding PRIVV vendor) are rolled up into the existing "Penn State Office of Physical Plant" vendor row if one exists, or appended as a new row if it does not.
- The comparison sorts results by the absolute value of the delta, largest discrepancies first.
- The summary bar at the top of the comparison window shows overall PRIVV total, eBuilder total, and net delta.


KNOWN LIMITATIONS

- The tool expects eBuilder exports to follow a consistent column order. If your export has a different structure, column mapping may be incorrect.
- Fuzzy matching uses a low threshold (40) which may occasionally produce incorrect matches. Review the "Match Note" column in the intermediate data if results look unexpected.
- The output file is always written to the directory where the script is located and is always named privv_importable.csv. Existing files with that name will be overwritten.
