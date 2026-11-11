# Excel Financial Summary Automation Tool

A lightweight, reliable, and formatting-preserving Excel automation tool built with Python and `openpyxl`.

## 📌 Project Overview
This tool automates the extraction of key financial summary cells (`Stock Value`, `MTF Debit`, `MTM`, `Booked PNL`, and `Date`) from a source Excel workbook containing multi-account sheets (`KSM`, `SDM`, and `KSM HUF`), and cleanly populates an existing Excel reference template without modifying or disturbing its fonts, alignments, column widths, row heights, merged cells, or layout.

---

## 🚀 Features
- **Strict Template Preservation**: Edits only the cells containing designated placeholders (`{PUTDATE}`, `{B3}`, `{B5}`, `{B9}`, `{B11}`, `{E3}`, etc.). Preserves all formatting, formulas, cell styles, and merged headers.
- **Accurate Cell Mapping**:
  - **Date**: Extracted from cell `A1`
  - **KSM**: `B3` (Stock Value), `B5` (MTF Debit), `B11` (Booked PNL), `B9` (MTM)
  - **SDM**: `E3` (Stock Value), `E5` (MTF Debit), `E11` (Booked PNL), `E9` (MTM)
  - **KSM HUF**: `H3` (Stock Value), `H5` (MTF Debit), `H11` (Booked PNL), `H9` (MTM)
- **Label Correction**: Automatically ensures headers remain `MTF Debit` (correcting typos like `NTF Debit`).
- **Data Integrity**: Preserves exact values and data types without recalculating or altering negative/positive sign formats. Ignores stock holdings tables and unrelated data.
- **Accidental Overwrite Protection**: Prevents overwriting the source or template files.
- **Interactive GUI**: Built with native `tkinter` featuring file browsers, smart path prefilling, live status reporting, and direct links to open generated files.
- **Command-Line Interface (CLI)**: Enables headless execution and script integration.

---

## 🛠️ Technology Stack
- **Python 3** (Standard Library: `tkinter`, `os`, `sys`, `re`, `datetime`, `unittest`)
- **`openpyxl`**: High-performance Excel reading and writing engine.
- No database, server, or cloud dependencies required.

---

## 📂 File Structure
```text
EXCEL_AUTOMATION/
├── excel_processor.py            # Core business logic and validation engine
├── gui.py                        # Modern Tkinter desktop graphical interface
├── main.py                       # Unified entry point (GUI or CLI mode)
├── test_automation.py            # Automated unit and integration test suite
├── template.xlsx                 # Reference template matching visual specifications
├── sample_source_23-10-2026.xlsx # Verified sample source workbook
├── processed_output.xlsx         # Generated output workbook
├── run.bat                       # One-click Windows batch launcher
├── requirements.txt              # Minimal dependencies
└── README.md                     # Documentation
```

---

## ⚡ Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Launch the Application (GUI)
Double-click `run.bat` or run:
```bash
python main.py
```

### 3. Run via Command-Line (CLI)
```bash
python main.py --source sample_source_23-10-2026.xlsx --template template.xlsx --output processed_output.xlsx
```

### 4. Run Automated Tests
```bash
python test_automation.py
```

---

## 📋 Field & Coordinate Reference Table

| Template Section | Financial Field | Source Cell | Template Placeholder |
|---|---|---|---|
| **Header** | Today Date | `A1` | `{PUTDATE}` |
| **KSM** | Stock Value | `B3` | `{B3}` |
| **KSM** | MTF Debit | `B5` | `{B5}` |
| **KSM** | Booked PNL | `B11` | `{B11}` |
| **KSM** | MTM | `B9` | `{B9}` |
| **SDM** | Stock Value | `E3` | `{E3}` |
| **SDM** | MTF Debit | `E5` | `{E5}` |
| **SDM** | Booked PNL | `E11` | `{E11}` |
| **SDM** | MTM | `E9` | `{E9}` |
| **KSM HUF** | Stock Value | `H3` | `{H3}` |
| **KSM HUF** | MTF Debit | `H5` | `{H5}` |
| **KSM HUF** | Booked PNL | `H11` | `{H11}` |
| **KSM HUF** | MTM | `H9` | `{H9}` |
