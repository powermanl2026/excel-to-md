# Excel-to-Markdown Skill

This skill converts Excel design documents to Markdown files, excluding specified sheets (default: '変更履歴'). It preserves the directory structure of the input Excel files in the output Markdown files.

## Installation

The skill is already installed if you have access to this directory. To use it, simply run:

```
/excel-to-md [options]
```

## Usage

See `references/usage-examples.md` for detailed examples.

## Customization

To change the default directories or excluded sheets, modify the default values in `scripts/convert_excel_to_md.py` or override them via command-line arguments.

## Requirements

- Python 3.x
- openpyxl library (install via `pip install openpyxl` if not already installed)

## License

This skill is provided as-is for use with Claude Code.