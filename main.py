"""
main.py
Main entry point for Excel Financial Summary Automation Tool.
Supports both graphical user interface (GUI) and command-line execution (CLI).
"""

import sys
import argparse
from excel_processor import process_excel_summary, ExcelAutomationError


def main():
    parser = argparse.ArgumentParser(
        description="Automate extraction of financial summary cells from source Excel workbook into a target template."
    )
    parser.add_argument("--source", "-s", help="Path to source Excel workbook (.xlsx)")
    parser.add_argument("--template", "-t", help="Path to template Excel workbook (.xlsx)")
    parser.add_argument("--output", "-o", help="Path for generated output Excel file (.xlsx)")
    parser.add_argument("--gui", "-g", action="store_true", help="Launch Tkinter Graphical User Interface")

    args = parser.parse_args()

    # If any file arguments are provided, run in CLI mode
    if args.source or args.template or args.output:
        if not (args.source and args.template and args.output):
            print("Error: When running in command-line mode, --source, --template, and --output are all required.")
            sys.exit(1)

        try:
            print("=" * 60)
            print("Excel Financial Summary Automation Engine")
            print("=" * 60)
            print(f"Source file:   {args.source}")
            print(f"Template file: {args.template}")
            print(f"Output file:   {args.output}")
            print("Processing...")

            result = process_excel_summary(args.source, args.template, args.output)

            print("\n[SUCCESS] Summary generation complete!")
            print(f"Sheet Name:        {result['source_sheet']}")
            print(f"Extracted Date:    {result['date']}")
            print(f"Replaced Cells:    {result['replacement_count']}")
            print(f"Output Saved To:   {result['output_file']}")
            print("\nExtracted Values:")
            for coord, val in result["extracted_cells"].items():
                print(f"  - {coord}: {val}")
            print("=" * 60)
            sys.exit(0)
        except ExcelAutomationError as e:
            print(f"\n[ERROR] Automation Error: {e}", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            print(f"\n[ERROR] Unexpected error: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        # Default behavior: launch GUI
        from gui import launch_gui
        launch_gui()


if __name__ == "__main__":
    main()
