EBuilder to PRIVV Processor
============================

**OVERVIEW**

This tool processes the EBuilder data and Privv data and sum's up indivudiual entries then compares them returning a sheet of information ranging from the commited amount to the delta(difference) between the EBuilder and Privv.


**REQUIREMENTS**

Install dependencies with:
  pip install -r script.txt

- Python 3.8 or higher
- The following Python packages:
  - pandas
  - rapidfuzz
  - tkinter (included with most Python installations)




**HOW TO RUN**


Run the program within either an IDE or terminal as such:
  python3 PXBComparer.py

A GUI window will open with the following controls.


**USING THE TOOL**


Step 1 - Select eBuilder Files
Click "Select eBuilder Files" to choose one or more exported eBuilder CSV files. Multiple files will be combined and deduplicated automatically. Click "Clear" to reset the selection.

Step 2 - Select PRIVV File
Click "Select PRIVV File" to choose your PRIVV reference CSV. This file must contain at minimum a "Vendor" column and an "Amount" column. If it also contains a "Code" column, those codes will be used to supplement the fuzzy matching logic.

Step 3 - Run Compare

Compare Vendors:
Opens a separate window showing a side-by-side dollar comparison of every vendor found in the PRIVV file against the matched totals in eBuilder. Rows are color coded: green for matches, yellow when eBuilder is higher, red when eBuilder is lower, and gray when no data exists on either side. You can filter by vendor name and export the comparison to CSV.


**VENDOR COMPARISON NOTES**


- The comparison sorts results by the absolute value of the delta, largest discrepancies first.
- The summary bar at the top of the comparison window shows overall PRIVV total, eBuilder total, and net delta.


**KNOWN LIMITATIONS**

- The tool expects eBuilder exports to follow a consistent column order. If your export has a different structure, column mapping may be incorrect.
- Fuzzy matching uses a low threshold (40) which may occasionally produce incorrect matches. Review the "Match Note" column in the intermediate data if results look unexpected.
