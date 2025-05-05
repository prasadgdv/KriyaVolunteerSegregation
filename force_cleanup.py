import os
import subprocess
import time
import glob

def force_cleanup():
    """
    More aggressive cleanup of Excel temporary files
    """
    print("Excel Force Cleanup Utility")
    print("===========================")
    
    # Step 1: Kill all Excel processes
    try:
        print("\nStep 1: Terminating Excel processes...")
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], 
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        print("Excel processes terminated.")
        time.sleep(3)  # Wait a bit longer to ensure processes are fully terminated
    except Exception as e:
        print(f"Error terminating Excel: {str(e)}")
    
    # Step 2: Use a more direct pattern match to find temp files
    print("\nStep 2: Finding and removing temporary files...")
    
    # Base directory path
    base_dir = os.path.dirname(os.path.abspath(__file__))
    vol_files_dir = os.path.join(base_dir, "volunteer_files")
    
    # Different patterns for temporary Excel files
    patterns = [
        os.path.join(vol_files_dir, "~$*.xlsx"),  # Standard Excel temp files
        os.path.join(vol_files_dir, "~*.tmp"),     # Other Excel temp files
        os.path.join(vol_files_dir, "*.tmp")       # General temp files
    ]
    
    deleted_count = 0
    
    # Try each pattern
    for pattern in patterns:
        for file_path in glob.glob(pattern):
            try:
                filename = os.path.basename(file_path)
                os.remove(file_path)
                print(f"Deleted: {filename}")
                deleted_count += 1
            except Exception as e:
                print(f"Could not delete {os.path.basename(file_path)}: {str(e)}")
    
    print(f"\nCleanup complete: {deleted_count} files removed.")
    
    # Step 3: Check what's left in the folder
    print("\nRemaining files in volunteer_files folder:")
    for filename in os.listdir(vol_files_dir):
        if filename.startswith("~") or filename.endswith(".tmp"):
            print(f"- {filename}")
    
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    force_cleanup()