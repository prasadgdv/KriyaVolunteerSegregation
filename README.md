# Volunteer Sheets Manager

A Python-based tool for managing volunteer sheets, including:

- Creating volunteer Excel sheets from master data
- Converting Excel files to PDF format
- Repairing corrupted Excel files
- Cleaning up temporary files

## Scripts

- `create_volunteer_sheets.py` - Create individual Excel files for volunteers from a master spreadsheet
- `convert_to_pdf.py` - Convert Excel files to PDF format
- `repair_excel_files.py` - Fix corrupted Excel files using master data
- `retry_failed_conversions.py` - Retry converting failed Excel files to PDF
- `cleanup_temp_files.py` - Remove temporary files created during processing
- `force_cleanup.py` - Force cleanup of all temporary files and processes

## Usage

Each script can be run independently from the command line:

```bash
python create_volunteer_sheets.py
python convert_to_pdf.py
python repair_excel_files.py
python retry_failed_conversions.py
python cleanup_temp_files.py
python force_cleanup.py
```

## Requirements

- Python 3.x
- pandas
- win32com (for Excel operations)
