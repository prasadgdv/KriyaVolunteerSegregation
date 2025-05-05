# Volunteer Sheets Generator

A complete solution for generating and managing volunteer sheets from Excel data.

## Overview

This toolkit provides a set of Python scripts to:

1. Generate individual Excel files for each volunteer from a master data file
2. Convert these Excel files to PDF format
3. Handle errors, cleanup temporary files, and repair corrupted files

## Requirements

- Windows OS
- Python 3.x
- Required Python packages:
  - openpyxl
  - pandas
  - pywin32 (for PDF conversion)

## Installation

1. Install Python from [python.org](https://python.org)
2. Install required packages:
   ```
   pip install openpyxl pandas pywin32
   ```

## Scripts

### 1. Create Volunteer Sheets (`create_volunteer_sheets.py`)

This script extracts volunteer data from a master Excel file and creates individual Excel files for each volunteer.

#### Features:

- Creates a folder structure based on tabs in the input Excel file
- Handles duplicate volunteer names by appending phone numbers
- Formats Excel files with proper headers, borders, and cell alignments
- Creates PDF folder structure for later conversion

#### Usage:

```
python create_volunteer_sheets.py
```

The script will:

1. Prompt you to select an Excel file from the current directory
2. Process the selected file to extract volunteer data
3. Create a folder structure based on tabs in the Excel file
4. Generate individual Excel files for each volunteer
5. Ask if you want to convert to PDF

### 2. Convert Excel to PDF (`convert_to_pdf.py`)

This script converts the generated Excel files to PDF format.

#### Features:

- Uses parallel processing for faster conversion
- Optimizes PDF layout for better printing
- Handles errors and creates a log of failed conversions
- District and mandal selection options

#### Usage:

```
python convert_to_pdf.py
```

The script will:

1. Show available district folders
2. Prompt you to select a district or process all districts
3. Show available mandal folders within the selected district
4. Prompt you to select mandals or process all mandals
5. Convert Excel files to PDF using parallel processing
6. Show progress and statistics on successful/failed conversions

#### Command Line Options:

- `--test`: Process only the first 2 files (for testing)
- `--manual`: Manually specify input and output folders

```
python convert_to_pdf.py --test
python convert_to_pdf.py --manual
```

### 3. Repair Excel Files (`repair_excel_files.py`)

This script repairs any corrupted Excel files from the volunteer sheets.

#### Usage:

```
python repair_excel_files.py
```

### 4. Retry Failed Conversions (`retry_failed_conversions.py`)

This script attempts to convert Excel files that failed in the initial conversion.

#### Usage:

```
python retry_failed_conversions.py
```

### 5. Clean Up Temporary Files (`cleanup_temp_files.py`)

This script removes temporary files created during processing.

#### Usage:

```
python cleanup_temp_files.py
```

### 6. Force Cleanup (`force_cleanup.py`)

This script forces the removal of Excel processes that may be blocking file operations.

#### Usage:

```
python force_cleanup.py
```

## Workflow

### Basic Workflow:

1. Run `create_volunteer_sheets.py` to generate Excel files from master data
2. Run `convert_to_pdf.py` to convert Excel files to PDF format

### Advanced Workflow:

1. Run `create_volunteer_sheets.py` to generate Excel files
2. Run `convert_to_pdf.py` to convert files to PDF
3. If any conversions fail, run `retry_failed_conversions.py`
4. If Excel files are corrupted, run `repair_excel_files.py`
5. After completion, run `cleanup_temp_files.py` to remove temporary files

## Folder Structure

The system creates the following folder structure:

```
volunteer_sheets/
├── [Master Excel Files].xlsx
├── excels_[district]/
│   ├── [tab1]/
│   │   ├── [volunteer1].xlsx
│   │   ├── [volunteer2].xlsx
│   │   └── ...
│   ├── [tab2]/
│   │   └── ...
│   └── ...
└── pdfs_[district]/
    ├── [tab1]/
    │   ├── [volunteer1].pdf
    │   ├── [volunteer2].pdf
    │   └── ...
    ├── [tab2]/
    │   └── ...
    └── ...
```

## Handling Duplicates

When multiple volunteer entries have the same name but different phone numbers, the system:

1. Groups entries by volunteer name and phone number
2. Creates a single file for each unique (name + phone) combination
3. Appends the phone number to the filename for volunteers with duplicate names

For example, if "John Doe" appears with two phone numbers (1234567890 and 9876543210), two files will be created:

- `John Doe_1234567890.xlsx`
- `John Doe_9876543210.xlsx`

## Troubleshooting

### Excel Files Not Converting to PDF

- Make sure Microsoft Excel is installed on your computer
- Run `force_cleanup.py` to close any hanging Excel processes
- Try converting a smaller batch with `convert_to_pdf.py --test`

### Corrupted Excel Files

- Run `repair_excel_files.py` to attempt to repair corrupted files
- Check file permissions and ensure Microsoft Excel can access the files

### Other Issues

- Make sure all required packages are installed
- Check that you have sufficient disk space
- Ensure you have the necessary permissions to create files in the directory

## License

This project is free to use for personal and commercial purposes.
