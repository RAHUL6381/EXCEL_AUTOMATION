"""
excel_processor.py
Core processing engine for Excel Financial Summary Automation Tool.
Extracts specific financial summary cells from a source Excel workbook
and populates an existing Excel template with strict formatting preservation.
"""

import os
import re
import datetime
from typing import Dict, Any, Tuple, Optional, List
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# Required financial summary cell coordinates mapped to their account section and label
REQUIRED_SOURCE_CELLS = {
    "A1": {"section": "Date", "field": "Date"},
    "B3": {"section": "KSM", "field": "Stock Value"},
    "B5": {"section": "KSM", "field": "MTF Debit"},
    "B9": {"section": "KSM", "field": "MTM"},
    "B11": {"section": "KSM", "field": "Booked PNL"},
    "E3": {"section": "SDM", "field": "Stock Value"},
    "E5": {"section": "SDM", "field": "MTF Debit"},
    "E9": {"section": "SDM", "field": "MTM"},
    "E11": {"section": "SDM", "field": "Booked PNL"},
    "H3": {"section": "KSM HUF", "field": "Stock Value"},
    "H5": {"section": "KSM HUF", "field": "MTF Debit"},
    "H9": {"section": "KSM HUF", "field": "MTM"},
    "H11": {"section": "KSM HUF", "field": "Booked PNL"},
}

# Supported placeholder regex: {PUTDATE}, {B3}, {B5}, etc., with optional inner whitespace
PLACEHOLDER_REGEX = re.compile(r"\{\s*(PUTDATE|[BEH](?:3|5|9|11))\s*\}", re.IGNORECASE)


class ExcelAutomationError(Exception):
    """Custom exception class for automation errors with user-friendly messages."""
    pass


def validate_file_paths(source_path: str, template_path: str, output_path: str) -> None:
    """
    Validate input file paths and ensure safety against overwriting.
    """
    if not source_path or not source_path.strip():
        raise ExcelAutomationError("Please select a source Excel workbook.")
    if not template_path or not template_path.strip():
        raise ExcelAutomationError("Please select a template Excel workbook.")
    if not output_path or not output_path.strip():
        raise ExcelAutomationError("Please specify an output Excel file path.")

    # Check source file existence
    if not os.path.exists(source_path):
        raise ExcelAutomationError(f"Source file not found:\n{source_path}")
    if not source_path.lower().endswith((".xlsx", ".xlsm")):
        raise ExcelAutomationError("Source file must be an Excel workbook (.xlsx or .xlsm).")

    # Check template file existence
    if not os.path.exists(template_path):
        raise ExcelAutomationError(f"Template file not found:\n{template_path}")
    if not template_path.lower().endswith((".xlsx", ".xlsm")):
        raise ExcelAutomationError("Template file must be an Excel workbook (.xlsx or .xlsm).")

    # Ensure output is .xlsx
    if not output_path.lower().endswith((".xlsx", ".xlsm")):
        raise ExcelAutomationError("Output file must have an Excel extension (.xlsx or .xlsm).")

    # Safety checks against accidental overwriting
    abs_src = os.path.abspath(source_path).lower()
    abs_tpl = os.path.abspath(template_path).lower()
    abs_out = os.path.abspath(output_path).lower()

    if abs_src == abs_out:
        raise ExcelAutomationError(
            "Accidental overwrite prevented: Output file cannot be the same as the source workbook.\n"
            "Please choose a different destination for the output."
        )

    if abs_tpl == abs_out:
        raise ExcelAutomationError(
            "Accidental overwrite prevented: Output file cannot be the same as the template workbook.\n"
            "Please choose a different destination for the output."
        )

    # Check output directory writability
    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.exists(out_dir):
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            raise ExcelAutomationError(f"Cannot create output directory:\n{out_dir}\nError: {e}")


def find_summary_sheet(wb: Workbook) -> Tuple[Worksheet, str]:
    """
    Identify the worksheet containing the financial summary table.
    Supports single summary sheet or 4-sheet workbooks.
    """
    # 1. Look for sheet named 'KSM & SDM' (exact or case-insensitive)
    for name in wb.sheetnames:
        clean_name = name.strip().lower()
        if clean_name in ["ksm & sdm", "ksm and sdm", "summary", "financial summary"]:
            return wb[name], name

    # 2. Inspect sheets to find the one with key summary markers (e.g. Stock Value at A3, D3, G3 or B3/E3/H3)
    for name in wb.sheetnames:
        ws = wb[name]
        try:
            # Check cell coordinates
            b3 = ws["B3"].value
            e3 = ws["E3"].value
            h3 = ws["H3"].value
            a3 = ws["A3"].value
            
            # If Stock Value label is at A3 or B3 has a value
            if (a3 and "stock value" in str(a3).lower()) or (b3 and e3 and h3):
                return ws, name
        except Exception:
            continue

    # 3. Fallback to active or first sheet if it has the required coordinates
    first_ws = wb.active or wb.worksheets[0]
    return first_ws, first_ws.title


def extract_financial_data(source_path: str) -> Dict[str, Any]:
    """
    Extract date from A1 and 12 financial summary cells from the source workbook.
    Preserves original data types and values without recalculation.
    """
    try:
        # Load workbook with data_only=False first to inspect
        wb_formulas = openpyxl.load_workbook(source_path, data_only=False)
        # Also load with data_only=True in case cached calculation values are needed
        wb_values = openpyxl.load_workbook(source_path, data_only=True)
    except Exception as e:
        raise ExcelAutomationError(f"Failed to open source Excel workbook:\n{e}")

    ws_f, sheet_name = find_summary_sheet(wb_formulas)
    ws_v = wb_values[sheet_name]

    extracted = {
        "source_sheet": sheet_name,
        "date_value": None,
        "date_format": None,
        "cells": {},
        "missing_cells": [],
    }

    # Extract date from A1
    date_cell_f = ws_f["A1"]
    date_cell_v = ws_v["A1"]
    raw_date = date_cell_v.value if date_cell_v.value is not None else date_cell_f.value

    if raw_date is None:
        extracted["missing_cells"].append("A1 (Date)")
    else:
        extracted["date_value"] = raw_date
        extracted["date_format"] = date_cell_f.number_format

    # Extract 12 financial cells
    financial_coords = ["B3", "B5", "B9", "B11", "E3", "E5", "E9", "E11", "H3", "H5", "H9", "H11"]
    for coord in financial_coords:
        meta = REQUIRED_SOURCE_CELLS[coord]
        cf = ws_f[coord]
        cv = ws_v[coord]

        val = cv.value if cv.value is not None else cf.value

        if val is None or (isinstance(val, str) and not val.strip()):
            extracted["missing_cells"].append(f"{coord} ({meta['section']} - {meta['field']})")
        else:
            extracted["cells"][coord] = {
                "value": val,
                "number_format": cf.number_format,
                "data_type": cf.data_type,
                "section": meta["section"],
                "field": meta["field"]
            }

    wb_formulas.close()
    wb_values.close()

    if extracted["missing_cells"]:
        missing_list = "\n - ".join(extracted["missing_cells"])
        raise ExcelAutomationError(
            f"The following required cells are missing or empty in source sheet '{sheet_name}':\n - {missing_list}\n\n"
            "Processing stopped to prevent generating an incomplete summary."
        )

    return extracted


def populate_template(template_path: str, output_path: str, extracted_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Populates the existing template workbook:
    - Replaces only designated placeholders: {PUTDATE}, {B3}, {B5}, {B9}, {B11}, {E3}, {E5}, {E9}, {E11}, {H3}, {H5}, {H9}, {H11}.
    - Corrects any accidental 'NTF DEBIT' typo in headers to 'MTF DEBIT' as required.
    - Preserves all template formatting, fonts, alignments, borders, merged cells, and layout.
    """
    try:
        wb = openpyxl.load_workbook(template_path)
    except Exception as e:
        raise ExcelAutomationError(f"Failed to open template workbook:\n{e}")

    date_val = extracted_data["date_value"]
    date_num_format = extracted_data.get("date_format")
    cells_data = extracted_data["cells"]

    # Track replacements for verification
    replacement_records = []
    found_placeholder_keys = set()

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None or not isinstance(cell.value, str):
                    continue

                cell_str = str(cell.value).strip()

                # Ensure 'MTF DEBIT' header is preserved if template had typo 'NTF DEBIT'
                if "NTF DEBIT" in cell_str.upper():
                    cell.value = re.sub(r"NTF\s+DEBIT", "MTF DEBIT", cell.value, flags=re.IGNORECASE)
                    continue

                # Check if cell contains or matches any placeholder
                # 1. Direct exact placeholder match (most common)
                upper_str = cell_str.upper()
                if upper_str in ["{PUTDATE}", "{PUTDATE }"] or re.fullmatch(r"\{\s*PUTDATE\s*\}", cell_str, re.IGNORECASE):
                    # Replace with date
                    if isinstance(date_val, datetime.datetime):
                        # If time is 00:00:00, use date object to prevent trailing time display
                        if date_val.hour == 0 and date_val.minute == 0 and date_val.second == 0:
                            cell.value = date_val.date()
                        else:
                            cell.value = date_val
                        if date_num_format and date_num_format.lower() != "general":
                            cell.number_format = date_num_format
                        else:
                            cell.number_format = "DD-MM-YYYY"
                    elif isinstance(date_val, datetime.date):
                        cell.value = date_val
                        cell.number_format = date_num_format or "DD-MM-YYYY"
                    else:
                        cell.value = str(date_val)

                    found_placeholder_keys.add("PUTDATE")
                    replacement_records.append({
                        "sheet": ws.title,
                        "cell": cell.coordinate,
                        "placeholder": "PUTDATE",
                        "value": str(cell.value)
                    })
                    continue

                coord_match = re.fullmatch(r"\{\s*([BEH](?:3|5|9|11))\s*\}", cell_str, re.IGNORECASE)
                if coord_match:
                    src_coord = coord_match.group(1).upper()
                    if src_coord in cells_data:
                        cell_info = cells_data[src_coord]
                        cell.value = cell_info["value"]
                        # Preserve number format if specified and not general
                        if cell_info.get("number_format") and cell_info["number_format"].lower() != "general":
                            cell.number_format = cell_info["number_format"]

                        found_placeholder_keys.add(src_coord)
                        replacement_records.append({
                            "sheet": ws.title,
                            "cell": cell.coordinate,
                            "placeholder": src_coord,
                            "value": str(cell.value)
                        })
                    continue

                # 2. Substring placeholder match (if placeholder is embedded in longer text)
                has_sub = False
                new_str = cell.value

                def sub_repl(match):
                    nonlocal has_sub
                    token = match.group(1).upper()
                    has_sub = True
                    found_placeholder_keys.add(token)
                    if token == "PUTDATE":
                        if isinstance(date_val, (datetime.datetime, datetime.date)):
                            return date_val.strftime("%d-%m-%Y")
                        return str(date_val)
                    elif token in cells_data:
                        return str(cells_data[token]["value"])
                    return match.group(0)

                new_str = PLACEHOLDER_REGEX.sub(sub_repl, new_str)
                if has_sub:
                    cell.value = new_str
                    replacement_records.append({
                        "sheet": ws.title,
                        "cell": cell.coordinate,
                        "placeholder": "EMBEDDED",
                        "value": str(cell.value)
                    })

    # Validate that placeholders were found in template
    expected_placeholders = {"PUTDATE", "B3", "B5", "B9", "B11", "E3", "E5", "E9", "E11", "H3", "H5", "H9", "H11"}
    missing_in_template = expected_placeholders - found_placeholder_keys

    if missing_in_template:
        missing_str = ", ".join(sorted(missing_in_template))
        wb.close()
        raise ExcelAutomationError(
            f"The selected template is missing the following required placeholders:\n{missing_str}\n\n"
            "Please ensure the template contains placeholders like {PUTDATE}, {B3}, {B5}, etc."
        )

    # Save output workbook
    try:
        wb.save(output_path)
    except PermissionError:
        wb.close()
        raise ExcelAutomationError(
            f"Permission denied when saving output file:\n{output_path}\n\n"
            "Please close the file if it is currently open in Excel and try again."
        )
    except Exception as e:
        wb.close()
        raise ExcelAutomationError(f"Failed to save output workbook:\n{e}")

    wb.close()
    return {
        "replacement_count": len(replacement_records),
        "records": replacement_records
    }


def verify_processed_output(output_path: str, extracted_data: Dict[str, Any]) -> None:
    """
    Verification check on the generated output workbook:
    - Confirms that NO unreplaced placeholders remain.
    - Confirms file size is valid and file opens cleanly.
    """
    try:
        wb = openpyxl.load_workbook(output_path)
    except Exception as e:
        raise ExcelAutomationError(f"Output verification failed: Unable to load saved file.\n{e}")

    unreplaced = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if PLACEHOLDER_REGEX.search(cell.value):
                        unreplaced.append(f"{ws.title}!{cell.coordinate}: {cell.value}")

    wb.close()

    if unreplaced:
        unrep_str = "\n - ".join(unreplaced[:10])
        raise ExcelAutomationError(
            f"Verification detected unreplaced placeholders in output workbook:\n - {unrep_str}\n"
            "Output file may be incomplete."
        )


def process_excel_summary(source_path: str, template_path: str, output_path: str) -> Dict[str, Any]:
    """
    Orchestrates the full end-to-end automation workflow:
    1. Validates inputs & prevents accidental overwriting.
    2. Extracts 12 financial summary cells & date from source workbook.
    3. Populates template while strictly preserving formatting.
    4. Validates output integrity.
    Returns a dictionary summarizing the operation results.
    """
    # Step 1: Validate paths
    validate_file_paths(source_path, template_path, output_path)

    # Step 2: Extract data
    extracted = extract_financial_data(source_path)

    # Step 3: Populate template
    result = populate_template(template_path, output_path, extracted)

    # Step 4: Verify output
    verify_processed_output(output_path, extracted)

    # Prepare user-friendly date string for summary
    date_val = extracted["date_value"]
    if isinstance(date_val, (datetime.datetime, datetime.date)):
        date_display = date_val.strftime("%d-%m-%Y")
    else:
        date_display = str(date_val)

    return {
        "success": True,
        "source_file": source_path,
        "template_file": template_path,
        "output_file": output_path,
        "source_sheet": extracted["source_sheet"],
        "date": date_display,
        "extracted_cells": {
            coord: info["value"] for coord, info in extracted["cells"].items()
        },
        "replacement_count": result["replacement_count"],
        "message": f"Successfully processed {result['replacement_count']} placeholder cells into {os.path.basename(output_path)}"
    }
