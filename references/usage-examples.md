# Usage Examples for excel-to-md Skill

## Basic Usage

Convert all Excel files in the default design directory to Markdown in the default docs directory:

```
/excel-to-md
```

This uses the default directories:
- Excel dir: `/Users/powermanl/powermanl-prj/yoho-check/設計書`
- Output dir: `/Users/powermanl/powermanl-prj/yoho-check/docs`

## Custom Directories

Convert Excel files from a custom directory to a custom output directory:

```
/excel-to-md --excel-dir /path/to/my/excel --output-dir /path/to/my/output
```

## Excluding Specific Sheets

By default, the skill excludes the sheet named "変更履歴". To exclude additional sheets or change the excluded sheets:

```
/excel-to-md --excel-dir /path/to/excel --output-dir /path/to/output --exclude-sheets "変更履歴" "テンプレート" "サマリー"
```

Or specify as a space-separated list:

```
/excel-to-md --excel-dir /path/to/excel --output-dir /path/to/output --exclude-sheets "変更履歴 テンプレート サマリー"
```

## Processing Subdirectories

The skill recursively processes all .xlsx files under the specified Excel directory, preserving the directory structure in the output.

For example, given the structure:
```
excel/
├── file1.xlsx
├── subdir/
│   └── file2.xlsx
```

The output will be:
```
output/
├── file1.md
└── subdir/
    └── file2.md
```

## Notes

- The skill only processes files with the `.xlsx` extension (case-insensitive).
- If all sheets in an Excel file are excluded, no output file is generated for that file.
- The skill uses openpyxl to read Excel files, so it supports .xlsx format (not the older .xls format).
- All cell values are converted to strings to preserve formatting as text.
- Merged cells are handled by assigning the value of the top-left cell to all cells in the merge.