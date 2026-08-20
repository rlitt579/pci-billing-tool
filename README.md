# PCI Billing Tool

**Professional billing tool for Extended Care**

![Version](https://img.shields.io/badge/version-3.0-blue)
![Python](https://img.shields.io/badge/python-3.10+-green)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Mac-lightgrey)

---

## Overview

A simple, clean billing tool designed for easy handoff. Two main functions:

| Tab | Purpose |
|-----|---------|
| **Generate Bills** | Process attendance file → billing summary + Veracross import CSV |
| **Combine Files** | Merge multiple Excel files → single import-ready CSV |

---

## Quick Start

### Download & Run

1. Download the [latest release](../../releases)
2. Extract the zip
3. **Windows:** Double-click `build.bat`
4. **Mac:** Run `./build_mac.sh` in Terminal
5. Find your app in the `dist/` folder

### Requirements

- Python 3.10+
- Dependencies (installed automatically by build script):
  ```
  pip install customtkinter pandas openpyxl pyinstaller
  ```

---

## Features

### Generate Bills Tab

1. Select your attendance Excel/CSV file
2. Enter school year and billing date
3. Click **Generate Bills**

**Output files** (saved in same folder as input):
| File | Description |
|------|-------------|
| `_SUMMARY.xlsx` | Billing summary with hours and charges |
| `_BOOKBILL.csv` | Ready for Veracross import |
| `_SKIPPED.csv` | Rows that were ignored (if any) |

**Required columns in your file:**
- `Student` (or `Person Name`)
- `Time Spent`
- `Person ID` (optional - for Veracross)
- `Notes` (optional - for ignore keywords)

---

### Combine Files Tab

1. Select a folder containing Excel files
2. Enter the sheet name to pull from
3. Enter school year and billing date
4. Click **Combine Files**

**Notes:**
- Blank rows (empty columns A or B) are automatically filtered out
- `school_year` and `item_date` columns are updated to your values

---

## Ignore Feature

Rows with these keywords in the **Notes** column are automatically skipped:

```
skip, ignore, no charge, waived, monthly
```

Skipped rows are logged in `_SKIPPED.csv` with the reason.

**To add more keywords**, edit `Config.IGNORE_KEYWORDS` in `pci_simple.py`

---

## Staff/Faculty Discount

> **Note:** The discount logic is built in but requires a code edit to enable.

| Rate Type | Amount |
|-----------|--------|
| Normal | $12.00/hour |
| Staff (50% off) | $6.00/hour |

### To Enable

1. Open `pci_simple.py`
2. Find `_generate_bills()` function
3. Add staff children list:

```python
staff_list = ["Last, First", "Baker Wiese, Josephine", ...]

summary_df, bookbill_df, stats, skipped_df = process_attendance_file(
    filepath, school_year, item_date,
    staff_children=staff_list  # Add this parameter
)
```

---

## Configuration

Edit the `Config` class at the top of `pci_simple.py`:

```python
class Config:
    HOURLY_RATE      = 12.0      # Dollars per hour
    STAFF_DISCOUNT   = 0.50      # 50% discount (0.50 = half off)
    CATALOG_ITEM_FK  = "2504"    # Veracross catalog item ID
    DESCRIPTION      = "EC After/Before School Care Drop-In"
    
    IGNORE_KEYWORDS  = ["skip", "ignore", "no charge", "waived", "monthly"]
```

---

## Build Instructions

### Windows

```batch
build.bat
```
Output: `dist\PCI_Billing_Tool.exe`

### Mac

```bash
chmod +x build_mac.sh
./build_mac.sh
```
Output: `dist/PCI_Billing_Tool.app`

---

## Error Logging

If errors occur, details are saved to:

| Platform | Location |
|----------|----------|
| Mac | `~/pci_billing_errors.log` |
| Windows | `C:\Users\[name]\pci_billing_errors.log` |

---

## File Structure

```
pci-billing-tool/
├── README.md          # This file
├── pci_simple.py      # Main application (edit Config here)
├── build.bat          # Windows build script
└── build_mac.sh       # Mac build script
```

---

## License

Internal use only.

---

## Support

For questions about the billing process, contact your department administrator.

For technical issues:
1. Check the error log (see Error Logging above)
2. Verify required columns exist in your data
3. Try rebuilding the app if it won't open
