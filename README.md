# EBuilder → PRIVV Processor (Weighted Classifier + Fuzzy Match)

## Overview
**EXPcomparerv11.py** is a sophisticated expense classification and vendor matching system that processes EBuilder commitment data and automatically assigns cost codes using a combination of weighted classification rules and fuzzy string matching algorithms.

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

## Installation
```bash
pip install pandas openpyxl rapidfuzz
```

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

### Step 5: Review Output
The tool generates `privv_importable.csv` with the following columns:
- **Code** - Cost code (e.g., 601, 708, 501)
- **Vendor** - Company/vendor name
- **Description** - Original description from EBuilder
- **PO Number** - (Currently empty, for future use)
- **Item Description** - Human-readable cost code label (e.g., "Signage", "Video Displays")
- **Date Committed** - Formatted date
- **Type** - "Change Order" or "Original"
- **Amount** - Current commitment amount

## Cost Code Mapping

The tool recognizes 60+ cost codes organized by category:

| Range | Category | Example Codes |
|-------|----------|---------------|
| 100-199 | Pre-Construction & Site | 102 (Feasibility), 107 (Environmental), 111 (Site Survey) |
| 200-299 | Financing & Insurance | 201 (Financing), 202 (Insurance), 204 (Builders Risk) |
| 300-399 | Project Management | 303 (PM), 308 (Testing), 310 (Commissioning) |
| 400-499 | Architecture & Design | 401 (A/E), 408 (Geotechnical), 416 (Technology) |
| 500-599 | Construction & Renovation | 501 (Pre-Const), 503.1 (West Side), 503.3 (Facility Work) |
| 600-699 | Furnishings & Operations | 601 (Signage), 603 (Equipment), 604 (Facility Ops) |
| 700-799 | Audio/Visual & Technology | 703 (Audio), 704 (Video), 712 (A/V), 713 (WIFI) |
| 800-899 | Fees & Contingency | 802 (Impact Fees) |
| 900-999 | Contingency | 901 (Owner Contingency) |

## Classification Algorithm

### Step 1: Weighted Rules (Priority-Based)
The system applies 100+ hardcoded rules with weight thresholds:
- **Weight 1000** - Highest priority (e.g., "Populous" always = 403)
- **Weight 500** - High priority (e.g., "Internal Commitment" = 418)
- **Weight 200** - Standard priority (most vendor/keyword rules)
- **Weight 100-150** - Lower priority (generic keywords like "signage")

Example rule structure:
```
{code: "403", weight: 1000, fields: ["vendor", "desc"], keywords: ["populous"], note: "..."}
```

### Step 2: Fuzzy Matching (Fallback)
If no rules match, the system uses fuzzy matching:
- Normalizes vendor names (lowercase, remove punctuation)
- Uses `RapidFuzz` token_sort ratio (80%+ threshold)
- Searches against normalized PRIVV lookup data
- Records match confidence in "Match Note"

### Step 3: Default Assignment
- If neither rules nor fuzzy matching succeeds: defaults to **"999"** (Unmatched)
- Unmatched rows are logged for manual review

## Troubleshooting

### Q: Many rows show code "999" (Unmatched)
**A:** Add more rules for those vendors or update the PRIVV lookup file. Review unmatched rows in the console log.

### Q: Fuzzy matching seems to miss obvious matches
**A:** Check vendor name normalization. Some special characters are stripped. Verify PRIVV file has correct vendor names.

### Q: PO grouping is combining unrelated items
**A:** This is by design—the tool aggregates all POs by (Vendor, Cost Code). If you need finer granularity, modify the grouping logic around line 1527.

### Q: Excel export has formatting issues
**A:** Ensure openpyxl is properly installed: `pip install --upgrade openpyxl`

## File Outputs

| File | Description |
|------|-------------|
| `privv_importable.csv` | Main output, ready for PRIVV import |
| `report_[timestamp].xlsx` | (Optional) Styled Excel report with highlighting |
| Console/Log | Real-time processing status and unmatched entries |

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

## Notes for Version 11
- **Date**: 06/24/2026
- **Current Focus**: Minimizing unaccounted-for rows (code 999)
- **Improvement Area**: Better keyword matching and rule coverage for edge cases

## Support
For issues or questions about specific rules, check the `note` field in `CLASSIFICATION_RULES` or review the console log output for detailed match information.

---


