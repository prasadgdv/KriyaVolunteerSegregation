import os
import sys
import glob
import time
import openpyxl
import pandas as pd
from pathlib import Path
from create_volunteer_sheets import repair_excel_file

def find_corrupted_excel_files(folder_path):
    """
    Find potentially corrupted Excel files in a folder
    
    Args:
        folder_path (str): Path to the folder containing Excel files
        
    Returns:
        list: List of potentially corrupted Excel files
    """
    corrupted_files = []
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    
    if not excel_files:
        print(f"No Excel files found in {folder_path}")
        return []
    
    print(f"Found {len(excel_files)} Excel files in {folder_path}")
    print("Checking for corrupted files...")
    
    for i, excel_file in enumerate(excel_files, 1):
        filename = os.path.basename(excel_file)
        try:
            # Check if Excel file can be opened with win32com (how convert_to_pdf tries to open it)
            # This is a light check using pandas, which is more forgiving with corrupted files
            df = pd.read_excel(excel_file, engine='openpyxl')
            file_size = os.path.getsize(excel_file)
            
            # Files with very small size might be corrupted or incomplete
            if file_size < 8000:  # Arbitrary threshold for minimal valid Excel file
                print(f"⚠️ Suspicious file size ({file_size} bytes): {filename}")
                corrupted_files.append(excel_file)
                continue
                
            # Check if we can open it with openpyxl as well (needed for repair)
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.active
                
            # If we can access cells, it's likely good
            test_value = ws['A1'].value
            wb.close()
            
        except Exception as e:
            print(f"❌ Error with file {filename}: {str(e)}")
            corrupted_files.append(excel_file)
    
    return corrupted_files

def extract_data_from_master(master_file, volunteer_name):
    """
    Extract data for a specific volunteer from the master Excel file
    
    Args:
        master_file (str): Path to the master Excel file
        volunteer_name (str): Name of the volunteer to extract data for
        
    Returns:
        list: List of data rows for the volunteer
    """
    data_found = []
    try:
        # Try to open master file with pandas for better error recovery
        all_data = pd.read_excel(master_file)
        
        # Look for volunteer name in the 8th column (index 7)
        for _, row in all_data.iterrows():
            if row.iloc[7] == volunteer_name:
                data_found.append(list(row))
        
        if data_found:
            print(f"Found {len(data_found)} records for {volunteer_name} in master file")
        else:
            print(f"No data found for {volunteer_name} in master file")
    
    except Exception as e:
        print(f"Error extracting data from master file: {str(e)}")
    
    return data_found

def repair_corrupted_files(master_file, folder_path):
    """
    Repair corrupted Excel files by recreating them from the master file
    
    Args:
        master_file (str): Path to the master Excel file
        folder_path (str): Path to the folder containing potentially corrupted Excel files
    """
    if not os.path.exists(master_file):
        print(f"Master file not found: {master_file}")
        return False
        
    corrupted_files = find_corrupted_excel_files(folder_path)
    
    if not corrupted_files:
        print("No corrupted files found!")
        return True
    
    print(f"\nFound {len(corrupted_files)} potentially corrupted files")
    
    for i, excel_file in enumerate(corrupted_files, 1):
        filename = os.path.basename(excel_file)
        volunteer_name = os.path.splitext(filename)[0]
        
        print(f"\n[{i}/{len(corrupted_files)}] Repairing: {filename}")
        
        # Extract data for this volunteer from master file
        volunteer_data = extract_data_from_master(master_file, volunteer_name)
        
        if not volunteer_data:
            print(f"Could not find data for {volunteer_name} in the master file, skipping...")
            continue
            
        # Repair the file using the extracted data
        success = repair_excel_file(excel_file, volunteer_data)
        
        if success:
            print(f"✓ Successfully repaired file for {volunteer_name}")
        else:
            print(f"✗ Failed to repair file for {volunteer_name}")
    
    return True

def main():
    # Set default paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excels_folder = os.path.join(base_dir, "excels")
    
    if not os.path.exists(excels_folder):
        print(f"'excels' folder not found at: {excels_folder}")
        return
    
    # Find all Excel files in the base directory (master files)
    master_files = [f for f in os.listdir(base_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
    
    if not master_files:
        print("No master Excel files found in the current directory.")
        return
    
    # Print menu for master Excel selection
    print("\nAvailable master Excel files:")
    for i, file in enumerate(master_files):
        print(f"{i+1}. {file}")
    
    try:
        master_selection = int(input("\nSelect a master file number: ")) - 1
        if 0 <= master_selection < len(master_files):
            master_file = os.path.join(base_dir, master_files[master_selection])
            print(f"Selected master file: {master_file}")
        else:
            print("Invalid selection.")
            return
    except ValueError:
        print("Invalid input. Please enter a number.")
        return
    
    # Get all subfolders in the excels folder
    subfolders = [f.path for f in os.scandir(excels_folder) if f.is_dir()]
    
    if not subfolders:
        print(f"No tab subfolders found in: {excels_folder}")
        return
    
    print(f"\nAvailable tab folders to check for corrupted files:")
    for i, folder in enumerate(subfolders):
        tab_name = os.path.basename(folder)
        print(f"{i+1}. {tab_name}")
    
    try:
        folder_selection = input("\nEnter folder number(s) to check (comma-separated) or 'all' for all folders: ")
        
        if folder_selection.lower() == 'all':
            folders_to_check = subfolders
        else:
            selections = [int(x.strip()) - 1 for x in folder_selection.split(',')]
            folders_to_check = [subfolders[i] for i in selections if 0 <= i < len(subfolders)]
    except:
        print("Invalid input.")
        return
    
    total_corrupted = 0
    
    for folder in folders_to_check:
        tab_name = os.path.basename(folder)
        print(f"\nChecking folder: {tab_name}")
        
        repair_corrupted_files(master_file, folder)
    
    print("\nAll selected folders have been processed!")

if __name__ == "__main__":
    print("Excel File Repair Tool")
    print("=====================")
    main()
    
    print("\nPress Enter to exit...")
    input()