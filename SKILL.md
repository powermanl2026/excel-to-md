---
name: excel-to-md
description: Convert Excel files to Markdown format with configurable options
version: 1.0.0
---
# Excel to Markdown Conversion Skill

Converts Excel workbooks to Markdown files, one per worksheet, with configurable input/output locations and sheet exclusions.

## Usage
```
/excel-to-md --excel-dir /path/to/excel/files --output-dir /path/to/markdown/output
```

## Options
- `--excel-dir`: Directory containing Excel files (required)
- `--output-dir`: Directory for Markdown output (required)
- `--exclude-sheets`: Comma-separated list of sheet names to exclude (optional, default: `変更履歴`)

## Examples
Convert all Excel files in `./data/excel` to Markdown in `./docs`:
```
/excel-to-md --excel-dir ./data/excel --output-dir ./docs
```

Exclude specific sheets:
```
/excel-to-md --excel-dir ./data/excel --output-dir ./docs --exclude-sheets "Template,Summary"
```