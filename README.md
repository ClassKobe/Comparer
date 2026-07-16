# EBuilder → PRIVV Processor (Weighted Classifier + Fuzzy Match)

## Overview
**EXPcomparer.py** is a sophisticated expense classification and vendor matching system that processes EBuilder commitment data and automatically assigns cost codes using a combination of weighted classification rules and fuzzy string matching algorithms.

## WARNING

DO NOT UPLOAD THE CREDENTIALS.JSON ONTO GIT, GOOGLE DRIVE OR ANY PUBLIC THING KEEP IT BETWEEN THE TEAM
-Sharing said file can have it disabled in the future as google loves to do that apparently which will cause a portion of the code to fail. If this for some reason happens a new key will need to be made and without Kobe here there is potentional for a the whole alias list o be remade.


## Purpose
This tool helps financial teams:
- **Categorize** expenses from EBuilder into standardized cost codes (101-901 range)
- **Match** vendor names with high accuracy using fuzzy matching
- **Aggregate** purchase orders by vendor and cost code
- **Export** processed data to a CSV file ready for PRIVV import
- **Identify** unmatched or problematic entries for manual review

## Key Features
✅ **Weighted Classification System** - Rules have priority weights to handle overlapping matches  
✅ **Fuzzy Matching** - Uses RapidFuzz to match vendor names despite typos/variations  
✅ **Multi-Field Scanning** - Searches vendor name, description, and commitment type fields  
✅ **Purchase Order Grouping** - Automatically aggregates POs by vendor and cost code  
✅ **Detailed Logging** - Provides transparent match notes showing why each row was classified  
✅ **Excel Export** - Creates styled Excel reports with color-coded classifications  
✅ **User-Friendly GUI** - Simple Tkinter interface for selecting files and running comparisons  

## System Requirements
- Python 3.7+
- Required packages:
  - `pandas` - Data manipulation
  - `openpyxl` - Excel file handling
  - `rapidfuzz` - Fuzzy string matching
  - `tkinter` - GUI (usually included with Python)
  - `gitpython` - GUI (usually included with Python)
 
  - WHEN INSTALLING PYTHON ADD PYTHON.EXE TO PATH ELSE YOU WILL MANUALLY HAVE TO ADD YOUR FILE TO PATH (THIS IS DIFFICULT IF YOU MESS UP UNINSTALL PYTHON AND TRY AGAIN)

## Installation
```bash
pip install pandas gitpython rapidfuzz openpyxl gitpython```

## How to Use

### Step 1: Prepare Input Files
You need two CSV files:

**EBuilder File** (e.g., `ebuilder_data.csv`)
- Must contain columns: Description, Company, Current Commitment, Date, Commitment Type, Vendor
- Rows with empty descriptions are automatically filtered out

**PRIVV File** (e.g., `privv_lookup.csv`)
- Must contain columns: Vendor, Code
- Used to build the lookup dictionary for fuzzy matching

### Step 2: Run the Application
```bash
python python_file.py
```

### Step 3: Select Files
1. Click **"Select eBuilder Files"** - Choose one or more CSV files containing commitment data
2. Click **"Select PRIVV File"** - Choose the CSV file with vendor-to-code mappings
3. Review selected files in the labels below each button

### Step 4: Run Comparison
1. Click **"Compare Vendors"**
2. The program will:
   - Load both files
   - Apply classification rules (weighted)
   - Perform fuzzy matching on unmatched items
   - Group purchase orders
   - Export results
3. Monitor progress in the output box

### Step 5: Run invoices 
1. Click **"invoices tab then the big green button"**
2. The program will:
   - Load both files
   - Apply classification rules (weighted)
   - Perform fuzzy matching on unmatched items
   - Group purchase orders
   - Export results

### Step 6: Vendor allias's 
1. Type in the name of the vendor
2. Type in the allias
  -From here the code will read the google sheet next time you run it
  -Allowing for the user to input allias's for vendors (Or if the code is being annoying you can input it however u please)

### Step 7: Review Output
The tool generates `privv_importable.csv` with the following columns:
- **Code** - Cost code (e.g., 601, 708, 501)
- **Vendor** - Company/vendor name
- **Description** - Original description from EBuilder
- **PO Number** - (Currently empty, for future use)
- **Item Description** - Human-readable cost code label (e.g., "Signage", "Video Displays")
- **Date Committed** - Formatted date
- **Type** - "Change Order" or "Original"
- **Amount** - Current commitment amount

## Advanced Features

### Excel Report Generation (Optional)
The code can generate a styled Excel report with:
- Color-coded cost codes (by category)
- Sorted by code
- Professional formatting (borders, fonts, alignment)
- Conditional highlighting for unmatched rows

### Custom Rule Addition
To add a new classification rule, edit the `CLASSIFICATION_RULES` list:
```python
{"code": "999", "weight": 200, "fields": ["vendor", "desc"], "keywords": ["your_keyword"], "note": "Your rule description"}
```

### Extending the Lookup
Update `ITEM_MAP` dictionary to add new cost codes or descriptions:
```python
ITEM_MAP = {
    "999": "Your New Cost Code Label",
    ...
}
```

## Support
For issues or questions about specific rules, check the `note` field in `CLASSIFICATION_RULES` or review the console log output for detailed match information.

---


