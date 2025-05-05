import os
import sys
import glob
import time
import pandas as pd
from pathlib import Path

def find_failed_list_files(base_dir):
    """
    Find all failed_list_*.xlsx files in the PDF output folders
    
    Args:
        base_dir (str): Base directory path
        
    Returns:
        list: List of failed list file paths
    """
    failed_list_files = []
    pdf_folder = os.path.join(base_dir, "pdfs")
    
    if not os.path.exists(pdf_folder):
        print(f"PDF folder not found: {pdf_folder}")
        return []
        
    # Get all subfolders in the PDF folder
    subfolders = [f.path for f in os.scandir(pdf_folder) if f.is_dir()]
    
    for folder in subfolders:
        # Find failed_list_*.xlsx files in the folder
        failed_lists = glob.glob(os.path.join(folder, "failed_list_*.xlsx"))
        failed_list_files.extend(failed_lists)
    
    return failed_list_files

def collect_failed_files(failed_list_file):
    """
    Extract the list of failed files from a failed_list_*.xlsx file
    
    Args:
        failed_list_file (str): Path to the failed list file
        
    Returns:
        tuple: (tab_name, list of failed filenames)
    """
    failed_files = []
    tab_name = os.path.basename(failed_list_file).replace("failed_list_", "").replace(".xlsx", "")
    
    try:
        # Read the failed list Excel file
        df = pd.read_excel(failed_list_file)
        
        # Get the list of failed filenames
        if 'File Name' in df.columns:
            # Convert all values to strings to avoid float issues
            failed_files = [str(file) if file is not None else "" for file in df['File Name'].tolist()]
        else:
            # Try to use the first column as filename, converting to strings
            failed_files = [str(file) if file is not None else "" for file in df.iloc[:, 0].tolist()]
    
    except Exception as e:
        print(f"Error reading failed list file {failed_list_file}: {str(e)}")
    
    # Filter out empty strings
    failed_files = [f for f in failed_files if f]
    
    return tab_name, failed_files

def convert_excel_to_pdf_single_file(excel_file, pdf_path):
    """
    Convert a single Excel file to PDF with optimized settings
    
    Args:
        excel_file (str): Path to the Excel file
        pdf_path (str): Path to save the PDF file
        
    Returns:
        bool: True if successful, False otherwise
    """
    print(f"Converting: {os.path.basename(excel_file)}...")
    temp_excel_file = None
    
    try:
        # Import the win32com module for Excel automation
        import win32com.client
        import numpy as np
        from win32com.client import constants
        
        # First try to open and fix the file with pandas
        try:
            print("Trying to fix mobile column values with pandas...")
            df = pd.read_excel(excel_file)
            
            # Find the mobile column
            mobile_col = None
            for col in df.columns:
                if 'mobile' in str(col).lower() or 'phone' in str(col).lower():
                    mobile_col = col
                    break
                    
            # If mobile column exists, fill all empty/NaN/#ERROR! values with "1111111111"
            modified = False
            if mobile_col:
                print(f"  Found mobile column: '{mobile_col}'")
                
                # Count issues before fixing
                empty_before = df[mobile_col].isna().sum()
                error_before = sum(1 for val in df[mobile_col] if isinstance(val, str) and (val == "#ERROR!" or val.startswith("#")))
                
                # Fix empty/NaN values
                df[mobile_col] = df[mobile_col].fillna("1111111111")
                
                # Fix #ERROR! values and other problematic values
                for i, val in enumerate(df[mobile_col]):
                    if pd.isna(val) or (isinstance(val, str) and (val == "#ERROR!" or val.startswith("#"))):
                        df.at[i, mobile_col] = "1111111111"
                        modified = True
                    elif val == "":
                        df.at[i, mobile_col] = "1111111111"
                        modified = True
                    # Try to clean mobile numbers if they contain invalid characters
                    elif isinstance(val, str):
                        cleaned_val = ""
                        for char in val:
                            if char.isdigit() or char in ['+', '-', ' ', '(', ')']:
                                cleaned_val += char
                        if cleaned_val == "":
                            df.at[i, mobile_col] = "1111111111"
                            modified = True
                        elif cleaned_val != val:
                            df.at[i, mobile_col] = cleaned_val
                            modified = True
                
                # Count issues after fixing
                empty_after = df[mobile_col].isna().sum()
                error_after = sum(1 for val in df[mobile_col] if isinstance(val, str) and (val == "#ERROR!" or val.startswith("#")))
                
                print(f"  Fixed {empty_before} empty values and {error_before} error values in mobile column")
                print(f"  Remaining empty: {empty_after}, Remaining errors: {error_after}")
                
            # If no mobile column found, create one
            else:
                print("  No mobile column found, adding one...")
                df["Mobile"] = "1111111111"
                modified = True
                print("  Added 'Mobile' column with default value '1111111111'")
            
            # If modifications were made, save the file with a temporary name
            if modified:
                temp_excel_file = excel_file + ".temp.xlsx"
                df.to_excel(temp_excel_file, index=False)
                print(f"  Saved fixed Excel file to {os.path.basename(temp_excel_file)}")
                # Use the fixed file for conversion
                excel_file_to_use = temp_excel_file
            else:
                print("  No modifications needed for mobile column")
                excel_file_to_use = excel_file
                
        except Exception as pandas_error:
            print(f"  Error fixing file with pandas: {str(pandas_error)}")
            print("  Continuing with original Excel file...")
            excel_file_to_use = excel_file
        
        # Initialize the Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            # Open the Excel file (either original or fixed)
            print(f"Opening Excel file: {os.path.basename(excel_file_to_use)}...")
            wb = excel.Workbooks.Open(
                os.path.abspath(excel_file_to_use),
                UpdateLinks=False,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                CorruptLoad=1  # xlNormalLoad (try normal load first)
            )
            
            # Set print settings for all sheets
            for sheet in wb.Worksheets:
                # Get row count to calculate optimal settings
                used_range = sheet.UsedRange
                last_row = used_range.Row + used_range.Rows.Count - 1
                
                # Set page orientation to PORTRAIT mode
                sheet.PageSetup.Orientation = 1  # xlPortrait (1=Portrait, 2=Landscape)
                
                # Set paper size to A4
                sheet.PageSetup.PaperSize = 9  # xlPaperA4
                
                # Set margins with significantly increased top margin for better spacing
                sheet.PageSetup.LeftMargin = 7.2    # 0.1 inches
                sheet.PageSetup.RightMargin = 7.2   # 0.1 inches
                sheet.PageSetup.TopMargin = 43.2    # 0.6 inches
                sheet.PageSetup.BottomMargin = 21.6 # 0.3 inches
                sheet.PageSetup.HeaderMargin = 7.2  # 0.1 inches
                sheet.PageSetup.FooterMargin = 7.2  # 0.1 inches
                
                # Handle scaling to fit more rows
                sheet.PageSetup.Zoom = False  # Don't use percentage zoom
                sheet.PageSetup.FitToPagesWide = 1  # Fit width on one page
                
                # Force fitting to pages tall if needed
                if last_row <= 45:  # For sheets with fewer rows
                    sheet.PageSetup.FitToPagesTall = 1
                else:
                    # Let it flow to multiple pages with better row density
                    sheet.PageSetup.FitToPagesTall = False
                    
                    # Calculate approximate rows per page
                    rows_per_page = 48  # Optimized target
                    pages_needed = (last_row + 2) / rows_per_page  # +2 for headers
                    pages_needed = max(1, round(pages_needed))
                    
                    # Adjust scaling for best fit
                    if pages_needed > 1:
                        sheet.PageSetup.FitToPagesTall = pages_needed
                
                # Center content horizontally only
                sheet.PageSetup.CenterHorizontally = True
                sheet.PageSetup.CenterVertically = False
                
                # Set print titles to repeat header rows
                sheet.PageSetup.PrintTitleRows = "$1:$2"
                
                # Turn off gridlines and row/column headings
                sheet.PageSetup.PrintGridlines = False
                sheet.PageSetup.PrintHeadings = False
            
            # Export to PDF with optimized settings
            try:
                wb.ExportAsFixedFormat(
                    Type=0,  # PDF format
                    Filename=os.path.abspath(pdf_path),
                    Quality=0,  # Standard quality
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                print(f"✓ Successfully created PDF")
                success = True
            except Exception as export_error:
                print(f"✗ Error during PDF export: {str(export_error)}")
                
                # Try a different approach - export each sheet separately
                try:
                    print("Trying alternate export method...")
                    temp_pdf_path = pdf_path + ".temp.pdf"
                    
                    # Try exporting just the first sheet
                    wb.Worksheets(1).ExportAsFixedFormat(
                        Type=0,
                        Filename=os.path.abspath(temp_pdf_path),
                        Quality=0,
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False
                    )
                    
                    # If successful, rename the temp file to the original filename
                    if os.path.exists(temp_pdf_path):
                        if os.path.exists(pdf_path):
                            os.remove(pdf_path)
                        os.rename(temp_pdf_path, pdf_path)
                        print(f"✓ Successfully created PDF using alternate method")
                        success = True
                    else:
                        success = False
                except Exception as alt_error:
                    print(f"✗ Alternative export method also failed: {str(alt_error)}")
                    success = False
            
            # Close the workbook without saving changes
            wb.Close(SaveChanges=False)
            
            return success
            
        except Exception as e:
            print(f"✗ Error converting {os.path.basename(excel_file)}: {str(e)}")
            # Try to close the workbook if it's still open
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
            return False
            
        finally:
            # Clean up: quit Excel
            try:
                excel.Quit()
            except:
                pass
            
            # Clean up temporary file if it exists
            if temp_excel_file and os.path.exists(temp_excel_file):
                try:
                    os.remove(temp_excel_file)
                    print(f"  Removed temporary Excel file")
                except:
                    pass
    
    except Exception as e:
        print(f"✗ Failed to initialize Excel: {str(e)}")
        
        # Clean up temporary file if it exists
        if temp_excel_file and os.path.exists(temp_excel_file):
            try:
                os.remove(temp_excel_file)
            except:
                pass
                
        return False

def update_failed_list(failed_list_file, still_failed_files):
    """
    Update the failed list file with files that still failed
    
    Args:
        failed_list_file (str): Path to the failed list file
        still_failed_files (list): List of files that still failed
    """
    if not still_failed_files:
        # If no files still failed, delete the failed list file
        try:
            os.remove(failed_list_file)
            print(f"All files successfully converted! Removed {os.path.basename(failed_list_file)}")
            return
        except Exception as e:
            print(f"Error removing {failed_list_file}: {str(e)}")
            return
    
    try:
        # Read the existing failed list
        df = pd.read_excel(failed_list_file)
        
        # Filter rows for files that are still failing
        if 'File Name' in df.columns:
            df_updated = df[df['File Name'].isin(still_failed_files)]
        else:
            # If columns don't match expected format, create a new dataframe
            df_updated = pd.DataFrame({
                'File Name': still_failed_files,
                'Error Message': ['Conversion failed after retry'] * len(still_failed_files),
                'Date/Time': [time.strftime("%Y-%m-%d %H:%M:%S")] * len(still_failed_files)
            })
        
        # Save the updated failed list
        df_updated.to_excel(failed_list_file, index=False)
        print(f"Updated failed list: {os.path.basename(failed_list_file)} with {len(still_failed_files)} remaining files")
        
    except Exception as e:
        print(f"Error updating failed list {failed_list_file}: {str(e)}")

def retry_failed_conversions():
    """
    Retry converting all failed Excel files to PDF
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Find all failed list files
    failed_list_files = find_failed_list_files(base_dir)
    
    if not failed_list_files:
        print("No failed conversion lists found.")
        return
    
    print(f"Found {len(failed_list_files)} failed conversion lists")
    
    # Process each failed list file
    total_files = 0
    total_success = 0
    total_still_failed = 0
    
    for failed_list_file in failed_list_files:
        # Get the tab name and list of failed files
        tab_name, failed_files = collect_failed_files(failed_list_file)
        
        if not failed_files:
            print(f"No failed files found in {os.path.basename(failed_list_file)}")
            continue
        
        print(f"\nProcessing failed files for tab: {tab_name}")
        print(f"Found {len(failed_files)} failed files to retry")
        
        # Get the source and destination folders
        source_folder = os.path.join(base_dir, "excels", tab_name.lower())
        dest_folder = os.path.join(base_dir, "pdfs", tab_name.lower())
        
        if not os.path.exists(source_folder):
            print(f"Source folder not found: {source_folder}")
            continue
            
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
            print(f"Created destination folder: {dest_folder}")
        
        still_failed_files = []
        success_count = 0
        
        # Retry converting each failed file
        for filename in failed_files:
            total_files += 1
            # Ensure filename is a string
            if not isinstance(filename, str):
                filename = str(filename)
            
            excel_file = os.path.join(source_folder, filename)
            pdf_name = os.path.splitext(filename)[0] + ".pdf"
            pdf_path = os.path.join(dest_folder, pdf_name)
            
            if not os.path.exists(excel_file):
                print(f"Excel file not found: {filename}")
                still_failed_files.append(filename)
                continue
            
            # Try to convert the file with extra retries
            success = False
            for attempt in range(3):  # Try up to 3 times
                if attempt > 0:
                    print(f"Retry attempt {attempt+1} for {filename}...")
                    time.sleep(2)  # Wait before retrying
                
                success = convert_excel_to_pdf_single_file(excel_file, pdf_path)
                if success:
                    break
            
            if success:
                success_count += 1
                total_success += 1
            else:
                still_failed_files.append(filename)
                total_still_failed += 1
        
        print(f"\nRetry results for {tab_name}:")
        print(f"  - {success_count} files successfully converted")
        print(f"  - {len(still_failed_files)} files still failing")
        
        # Update the failed list file
        update_failed_list(failed_list_file, still_failed_files)
    
    print(f"\nOverall retry results:")
    print(f"  - {total_files} total files processed")
    print(f"  - {total_success} files successfully converted")
    print(f"  - {total_still_failed} files still failing")
    
    if total_still_failed > 0:
        print("\nFor files that still fail, you can try using repair_excel_files.py to fix them.")

def convert_single_file_test():
    """
    Convert a single file to test the conversion process
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    print("Single File Conversion Test")
    print("==========================")
    
    # Ask user which tab to process
    pdf_folder = os.path.join(base_dir, "pdfs")
    subfolders = [f.name for f in os.scandir(pdf_folder) if f.is_dir()]
    
    print("Available tabs:")
    for i, folder in enumerate(subfolders, 1):
        print(f"{i}. {folder}")
    
    folder_choice = input("Enter tab number to check: ")
    try:
        folder_index = int(folder_choice) - 1
        tab_name = subfolders[folder_index]
    except (ValueError, IndexError):
        print("Invalid choice. Exiting.")
        return
    
    # Get a list of Excel files in the chosen tab
    source_folder = os.path.join(base_dir, "excels", tab_name)
    if not os.path.exists(source_folder):
        print(f"Source folder not found: {source_folder}")
        return
    
    excel_files = [f.name for f in os.scandir(source_folder) if f.is_file() and f.name.endswith('.xlsx') and not f.name.startswith('~$')]
    
    # Show the Excel files to the user
    print(f"\nAvailable Excel files in {tab_name}:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file}")
    
    file_choice = input("Enter file number to convert: ")
    try:
        file_index = int(file_choice) - 1
        filename = excel_files[file_index]
    except (ValueError, IndexError):
        print("Invalid choice. Exiting.")
        return
    
    # Set up the paths
    excel_file = os.path.join(source_folder, filename)
    dest_folder = os.path.join(base_dir, "pdfs", tab_name)
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    
    pdf_name = os.path.splitext(filename)[0] + ".pdf"
    pdf_path = os.path.join(dest_folder, pdf_name)
    
    # Try to convert the file
    print(f"\nAttempting to convert {filename} to PDF...")
    success = convert_excel_to_pdf_single_file(excel_file, pdf_path)
    
    if success:
        print(f"\n✓ Successfully converted {filename} to PDF!")
        print(f"PDF saved to: {pdf_path}")
    else:
        print(f"\n✗ Failed to convert {filename} to PDF.")

def check_failed_excel_file(excel_file):
    """
    Check what's wrong with a failed Excel file, particularly the mobile column
    
    Args:
        excel_file (str): Path to the Excel file to check
        
    Returns:
        bool: True if issues were found and reported, False otherwise
    """
    print(f"Examining file: {os.path.basename(excel_file)}")
    print("====================================")
    
    try:
        # First, try reading the file with pandas to check if it's readable
        try:
            df = pd.read_excel(excel_file)
            print("✓ File can be read with pandas")
            
            # Check for mobile column
            mobile_col = None
            for col in df.columns:
                if 'mobile' in str(col).lower() or 'phone' in str(col).lower():
                    mobile_col = col
                    break
            
            if mobile_col:
                print(f"Found mobile column: '{mobile_col}'")
                
                # Analyze mobile column values
                total_rows = len(df)
                empty_cells = df[mobile_col].isna().sum()
                error_cells = sum(1 for val in df[mobile_col] if isinstance(val, str) and (val == "#ERROR!" or val.startswith("#")))
                
                print(f"Total rows: {total_rows}")
                print(f"Empty cells in mobile column: {empty_cells} ({empty_cells/total_rows:.1%})")
                print(f"Error cells (#ERROR!, etc.): {error_cells} ({error_cells/total_rows:.1%})")
                
                # Show some examples of problematic values
                if error_cells > 0:
                    error_examples = [val for val in df[mobile_col] if isinstance(val, str) and (val == "#ERROR!" or val.startswith("#"))]
                    print("Examples of error values:")
                    for i, ex in enumerate(error_examples[:5]):
                        print(f"  {i+1}. {ex}")
                    if len(error_examples) > 5:
                        print(f"  ... and {len(error_examples)-5} more")
                
                # Check for non-standard mobile values (not numeric)
                non_standard = []
                for val in df[mobile_col]:
                    if pd.notna(val) and not (isinstance(val, str) and val.startswith("#")):
                        # Try to convert to string and check if it has non-numeric chars (except for standard phone chars)
                        str_val = str(val)
                        has_invalid = False
                        for char in str_val:
                            if not (char.isdigit() or char in ['+', '-', ' ', '(', ')']):
                                has_invalid = True
                                break
                        if has_invalid:
                            non_standard.append(str_val)
                
                if non_standard:
                    print(f"Found {len(non_standard)} non-standard mobile values:")
                    for i, val in enumerate(non_standard[:5]):
                        print(f"  {i+1}. {val}")
                    if len(non_standard) > 5:
                        print(f"  ... and {len(non_standard)-5} more")
            else:
                print("❌ No mobile/phone column found")
                
            # Check for other common issues
            print("\nChecking for other common issues:")
            for col in df.columns:
                col_errors = sum(1 for val in df[col] if isinstance(val, str) and (val == "#ERROR!" or val.startswith("#")))
                if col_errors > 0:
                    print(f"  - Column '{col}' has {col_errors} error values")
            
        except Exception as pd_error:
            print(f"❌ Cannot read file with pandas: {str(pd_error)}")
            
        # Then try with win32com to get more detailed Excel issues
        print("\nTrying to analyze with Excel API...")
        import win32com.client
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            wb = excel.Workbooks.Open(
                os.path.abspath(excel_file),
                UpdateLinks=False,
                ReadOnly=True,
                CorruptLoad=1
            )
            
            print("✓ File can be opened with Excel API")
            
            # Check for errors in each sheet
            for sheet_idx in range(1, wb.Sheets.Count + 1):
                sheet = wb.Sheets(sheet_idx)
                sheet_name = sheet.Name
                used_range = sheet.UsedRange
                print(f"\nAnalyzing Sheet: {sheet_name}")
                
                # Find the mobile column
                mobile_col_idx = None
                mobile_col_name = ""
                
                for col in range(1, used_range.Columns.Count + 1):
                    cell_value = sheet.Cells(1, col).Value
                    if cell_value and isinstance(cell_value, str) and ("mobile" in cell_value.lower() or "phone" in cell_value.lower()):
                        mobile_col_idx = col
                        mobile_col_name = cell_value
                        break
                
                if mobile_col_idx:
                    print(f"Found mobile column: '{mobile_col_name}' (Column {mobile_col_idx})")
                    
                    # Check cell values in the mobile column
                    error_count = 0
                    empty_count = 0
                    non_standard_count = 0
                    examples = []
                    
                    for row in range(2, used_range.Rows.Count + 1):  # Start from row 2 (data rows)
                        cell = sheet.Cells(row, mobile_col_idx)
                        cell_value = cell.Value
                        
                        if cell_value is None:
                            empty_count += 1
                        elif isinstance(cell_value, str):
                            if cell_value == "#ERROR!" or cell_value.startswith("#"):
                                error_count += 1
                                if len(examples) < 5:
                                    examples.append(f"Row {row}: {cell_value}")
                            else:
                                # Check for non-standard characters
                                has_invalid = False
                                for char in cell_value:
                                    if not (char.isdigit() or char in ['+', '-', ' ', '(', ')']):
                                        has_invalid = True
                                        break
                                if has_invalid and len(examples) < 5:
                                    non_standard_count += 1
                                    examples.append(f"Row {row}: {cell_value}")
                        elif not isinstance(cell_value, (int, float)):
                            non_standard_count += 1
                            if len(examples) < 5:
                                examples.append(f"Row {row}: {str(cell_value)}")
                    
                    total_rows = used_range.Rows.Count - 1  # Excluding header
                    print(f"Total rows: {total_rows}")
                    print(f"Empty cells: {empty_count} ({empty_count/total_rows:.1%})")
                    print(f"Error values: {error_count} ({error_count/total_rows:.1%})")
                    print(f"Non-standard values: {non_standard_count} ({non_standard_count/total_rows:.1%})")
                    
                    if examples:
                        print("Examples of problematic values:")
                        for ex in examples:
                            print(f"  - {ex}")
                    
                    print("\nSuggested fix: Use updated converter that ignores errors and empty values in the mobile column")
            
            wb.Close(SaveChanges=False)
        except Exception as excel_error:
            print(f"❌ Error analyzing with Excel API: {str(excel_error)}")
        finally:
            excel.Quit()
        
        return True
    except Exception as e:
        print(f"❌ Error during analysis: {str(e)}")
        return False

def select_failed_file_to_check():
    """
    Select a failed file from a failed list to check what's wrong with it
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    print("Check Failed Excel File")
    print("======================")
    
    # Find all failed list files
    failed_list_files = find_failed_list_files(base_dir)
    
    if not failed_list_files:
        print("No failed conversion lists found.")
        return
    
    print("Failed conversion lists:")
    for i, file_path in enumerate(failed_list_files, 1):
        tab_name = os.path.basename(file_path).replace("failed_list_", "").replace(".xlsx", "")
        print(f"{i}. {tab_name}")
    
    list_choice = input("Enter the number of the failed list to check: ")
    try:
        list_index = int(list_choice) - 1
        failed_list_file = failed_list_files[list_index]
    except (ValueError, IndexError):
        print("Invalid choice. Exiting.")
        return
    
    # Get the tab name and failed files from the selected list
    tab_name, failed_files = collect_failed_files(failed_list_file)
    
    if not failed_files:
        print(f"No failed files found in {os.path.basename(failed_list_file)}")
        return
    
    print(f"\nFailed files for tab: {tab_name}")
    for i, filename in enumerate(failed_files, 1):
        print(f"{i}. {filename}")
    
    file_choice = input("Enter file number to check: ")
    try:
        file_index = int(file_choice) - 1
        filename = failed_files[file_index]
    except (ValueError, IndexError):
        print("Invalid choice. Exiting.")
        return
    
    # Set up the path to the Excel file
    source_folder = os.path.join(base_dir, "excels", tab_name.lower())
    excel_file = os.path.join(source_folder, filename)
    
    if not os.path.exists(excel_file):
        print(f"Excel file not found: {excel_file}")
        return
    
    # Check what's wrong with the file
    check_failed_excel_file(excel_file)

if __name__ == "__main__":
    print("Retry Failed PDF Conversions")
    print("===========================")
    
    # Check if user wants to test a single file
    print("1. Test a single file conversion")
    print("2. Retry all failed conversions")
    print("3. Check what's wrong with a failed file")
    choice = input("Enter your choice (1/2/3): ")
    
    if choice == "1":
        convert_single_file_test()
    elif choice == "3":
        select_failed_file_to_check()
    else:
        retry_failed_conversions()
    
    print("\nPress Enter to exit...")
    input()