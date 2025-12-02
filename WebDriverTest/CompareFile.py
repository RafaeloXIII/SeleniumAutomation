import pandas as pd
import os

def get_latest_file(folder_path):
    if not os.path.exists(folder_path):
        print(f"Folder '{folder_path}' does not exist.")
        return None

    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not files:
        print(f"No .xlsx files found in '{folder_path}'.")
        return None

    latest_file = max(files, key=lambda f: os.path.getctime(os.path.join(folder_path, f)))
    return os.path.join(folder_path, latest_file)

def compare_columns(new_file, old_file):
    new_df = pd.read_excel(new_file, engine='openpyxl')
    old_df = pd.read_excel(old_file, engine='openpyxl')

    new_columns = set(new_df.columns)
    old_columns = set(old_df.columns)

    added_columns = new_columns - old_columns
    removed_columns = old_columns - new_columns

    if added_columns or removed_columns:
        print(f"Changes detected in columns between {new_file} and {old_file}:")
        if added_columns:
            print("Added columns:", added_columns)
        if removed_columns:
            print("Removed columns:", removed_columns)
    else:
        print(f"No changes detected in columns between {new_file} and {old_file}.")

def check_current_folders(current_folders, history_folders):
    if len(current_folders) != len(history_folders):
        print("Number of current folders and history folders should be the same.")
        return

    for i, current_folder in enumerate(current_folders):
        latest_recent_file = get_latest_file(current_folder)
        latest_history_file = get_latest_file(history_folders[i])

        if latest_recent_file and latest_history_file:
            compare_columns(latest_recent_file, latest_history_file)
        else:
            print(f"Could not find files to compare in {current_folder} or {history_folders[i]}.")

# Paths to folders
current_folders = [r"C:\Users\Rafael Braga\Documents\Compare\1"]
history_folders = [r"C:\Users\Rafael Braga\Documents\Compare\2"]

check_current_folders(current_folders, history_folders)
