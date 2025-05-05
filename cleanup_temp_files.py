import os
import shutil
import subprocess
import time

def kill_excel_processes():
    """
    Force close all Excel processes
    """
    try:
        print("Attempting to close Excel instances...")
        # Try to gracefully close Excel using taskkill
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        print("Excel processes closed.")
        # Give some time for the processes to fully terminate
        time.sleep(2)
        return True
    except Exception as e:
        print(f"Could not close Excel: {str(e)}")
        return False

def cleanup_temp_files(folder_path):
    """
    Clean up temporary Excel files in the specified folder
    
    Args:
        folder_path: Path to the folder to clean
    """
    count = 0
    
    # Check if the folder exists
    if not os.path.exists(folder_path):
        print(f"Folder not found: {folder_path}")
        return
        
    print(f"Scanning for temporary files in {folder_path}...")
    
    # List all files in the directory
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
        # Check if it's a temporary Excel file (starts with ~$)
        if filename.startswith("~$") and os.path.isfile(file_path):
            try:
                os.remove(file_path)
                print(f"Deleted: {filename}")
                count += 1
            except Exception as e:
                print(f"Could not delete {filename}: {str(e)}")
    
    print(f"\nCleanup complete: {count} temporary files removed.")

if __name__ == "__main__":
    # Path to the volunteer files folder
    volunteer_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "volunteer_files")
    
    print("Excel Temporary File Cleanup Utility")
    print("===================================")
    
    # First kill any Excel processes
    kill_excel_processes()
    
    # Then clean up the volunteer files folder
    cleanup_temp_files(volunteer_folder)
    
    print("\nPress Enter to exit...")
    input()