import os
import glob
import time
import subprocess

# Lock file to prevent re-running
LOCK_FILE = "automation_lock.txt"

def get_most_recent_file(folder_path):
    try:
        files = glob.glob(os.path.join(folder_path, '*'))
        if not files:
            return None
        most_recent_file = max(files, key=os.path.getmtime)
        return most_recent_file
    except Exception as e:
        print(f"Error getting most recent file: {e}")
        return None

def is_xlsx(file_path):
    return file_path.endswith('.xlsx')

def run_automation(automation_path):
    try:
        print("Running the automation again...")
        subprocess.run(['python', automation_path], check=True)
        print("Automation script completed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error running automation: {e}")
    except FileNotFoundError:
        print(f"Automation script not found at {automation_path}. Please check the path.")
    except Exception as e:
        print(f"An unexpected error occurred while running automation: {e}")

def create_lock_file(folder_path):
    with open(os.path.join(folder_path, LOCK_FILE), 'w') as lock_file:
        lock_file.write('locked')

def remove_lock_file(folder_path):
    try:
        os.remove(os.path.join(folder_path, LOCK_FILE))
    except FileNotFoundError:
        pass

def check_and_run_automation(download_folder, automation_path):
    try:
        # Check for the lock file to prevent recursion
        if os.path.exists(os.path.join(download_folder, LOCK_FILE)):
            print("Automation is already running")
            return

        # Create a lock file to prevent another automation run
        create_lock_file(download_folder)

        # Get most recent file in the folder
        most_recent_file = get_most_recent_file(download_folder)

        if most_recent_file:
            print(f"Most recent file: {most_recent_file}")
            # If the file is not an xlsx, run the automation again
            if not is_xlsx(most_recent_file):
                print("Running the automation again")
                run_automation(automation_path)
            else:
                print("No action needed")
        else:
            print("No files found")
        
        # Remove the lock file once the automation is complete
        remove_lock_file(download_folder)
    except Exception as e:
        print(f"Error in check_and_run_automation: {e}")
        remove_lock_file(download_folder)