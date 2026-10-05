#!/usr/bin/env python3
"""
Convert Excel design documents to Markdown files.
Excludes specified sheets (default: '変更履歴').
Outputs to a specified directory mirroring the original structure under the Excel root.
"""

import os
import sys
import argparse
from pathlib import Path
from openpyxl import load_workbook

def sheet_to_markdown(ws):
    """
    Convert a worksheet to a Markdown string.
    Returns the Markdown text.
    """
    # Determine dimensions
    max_row = ws.max_row
    max_column = ws.max_column

    # If the sheet is empty, return empty string
    if max_row == 0 or max_column == 0:
        return ""

    # Create a 2D list for values, initialized with None
    values = [[None for _ in range(max_column)] for _ in range(max_row)]

    # Fill values from cells
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_column):
        for cell in row:
            # cell.row and cell.column are 1-indexed
            values[cell.row - 1][cell.column - 1] = cell.value

    # Apply merged cells: for each merged range, set all cells to the value of the top-left cell
    for merged_range in ws.merged_cells.ranges:
        # Get top-left cell's value
        top_left = ws[merged_range.start_cell.coordinate]
        fill_value = top_left.value
        # Iterate over rows and columns in the range
        for row_idx in range(merged_range.min_row, merged_range.max_row + 1):
            for col_idx in range(merged_range.min_col, merged_range.max_col + 1):
                # Adjust to 0-index
                values[row_idx - 1][col_idx - 1] = fill_value

    # Convert to strings, replacing None with empty string
    str_values = [[("" if v is None else str(v)) for v in row] for row in values]

    # Build Markdown table
    lines = []
    # Header row (first row)
    header = str_values[0]
    lines.append("| " + " | ".join(header) + " |")
    # Separator row
    separator = ["---"] * len(header)
    lines.append("| " + " | ".join(separator) + " |")
    # Data rows (remaining rows)
    for row in str_values[1:]:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


def process_excel_file(excel_path, output_root, exclude_sheets):
    """
    Convert a single Excel file to Markdown.
    """
    # Compute relative path from the Excel root (we need to pass the excel root as well)
    # We'll change the function signature to also take excel_root to compute relative path.
    # Actually, we need to know the excel root to mirror the structure.
    # We'll change the main loop to pass excel_root and output_root.
    pass

def main():
    parser = argparse.ArgumentParser(
        description='Convert Excel files to Markdown files, excluding specified sheets.'
    )
    parser.add_argument(
        '--excel-dir',
        type=str,
        default='/Users/powermanl/powermanl-prj/yoho-check/設計書',
        help='Root directory containing Excel files (default: /Users/powermanl/powermanl-prj/yoho-check/設計書)'
    )
    parser.add_argument(
        '--output-dir',
        type=str,
        default='/Users/powermanl/powermanl-prj/yoho-check/docs',
        help='Root directory for Markdown output (default: /Users/powermanl/powermanl-prj/yoho-check/docs)'
    )
    parser.add_argument(
        '--exclude-sheets',
        type=str,
        nargs='*',
        default=['変更履歴'],
        help='Sheet names to exclude (default: 変更履歴). Can specify multiple times or as space-separated list.'
    )

    args = parser.parse_args()

    excel_root = Path(args.excel_dir)
    output_root = Path(args.output_dir)
    exclude_sheets = set(args.exclude_sheets)

    # Walk the Excel root
    for root, dirs, files in os.walk(excel_root):
        for file in files:
            if file.lower().endswith(".xlsx"):
                excel_path = Path(root) / file
                # Compute relative path from excel_root
                rel_path = excel_path.relative_to(excel_root)
                # Change extension to .md
                md_rel_path = rel_path.with_suffix(".md")
                output_path = output_root / md_rel_path

                # Ensure output directory exists
                output_path.parent.mkdir(parents=True, exist_ok=True)

                # Load workbook
                wb = load_workbook(excel_path, data_only=True)

                all_md = []
                for sheet_name in wb.sheetnames:
                    if sheet_name in exclude_sheets:
                        continue
                    ws = wb[sheet_name]
                    # Optionally, we can add a sheet title as a heading
                    md = f"## シート: {sheet_name}\n\n"
                    md += sheet_to_markdown(ws)
                    all_md.append(md)

                wb.close()

                if not all_md:
                    # If all sheets were excluded, skip
                    continue

                # Combine: if multiple sheets, we separate with a newline
                full_md = "\n\n---\n\n".join(all_md)

                # Write to file
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(full_md)

                print(f"Converted: {excel_path} -> {output_path}")

if __name__ == "__main__":
    main()