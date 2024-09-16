import hashlib
import os
import sys
import time
from datetime import datetime

import pandas as pd


def tm(func):
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print(f"Execution time of {func.__name__}: {end_time - start_time:.2f} seconds")
        return result

    return wrapper


def file_checksum(file_path):
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
    except IOError:
        print(f"Can't open: {file_path}")
        return None
    return hash_md5.hexdigest()


@tm
def find_duplicates(folder_path):
    if not os.path.isdir(folder_path):
        print(f"Folder {folder_path} not found...")
        return

    data = []
    total_files = sum([len(files) for r, d, files in os.walk(folder_path)])
    processed_files = 0

    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            if len(file_path) >= 260:
                file_path = f"\\\\?\\{file_path}"
            try:
                file_size = os.path.getsize(file_path)
            except FileNotFoundError:
                print(f"File not found: {file_path}")
                continue
            file_ext = os.path.splitext(file)[1]
            data.append([file_size, file_ext, file_path, file])

            processed_files += 1
            progress = (processed_files / total_files) * 100
            sys.stdout.write(f"\rProgress: {progress:.2f} %")
            sys.stdout.flush()

    df = pd.DataFrame(data, columns=['Size', 'Extension', 'File Path', 'File Name'])

    duplicates = df[df.duplicated(subset=['Size', 'Extension'], keep=False)].copy()

    # Add a unique number to each group
    duplicates['Group_Number'] = duplicates.groupby(['Size', 'Extension']).ngroup()

    # Mark the file with the shortest name as 'master'
    duplicates['Master'] = duplicates.groupby(['Group_Number'])['File Name'].transform(
        lambda x: x == x.loc[x.str.len().idxmin()])
    print("\nProcessing files' checksum...")

    # Calculate checksums in parallel
    duplicates['Check_Sum'] = [file_checksum(file_path) for file_path in duplicates['File Path']]

    duplicates.sort_values(by=['Group_Number'], inplace=True, ascending=False)

    print("\nDuplicates search complete\n")

    return duplicates





def save_to_excel(duplicates_to_save, output_folder):
    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_file = os.path.join(output_folder, f"duplicates_{now}.xlsx")
    try:
        duplicates_to_save.to_excel(output_file, index=False)
    except:
        print("Can't save file in selected path. Enter another path")
        new_path=input()
        now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if not os.path.isdir(new_path):
            print(f"Folder {new_path} not found...")
            return
        output_file = os.path.join(output_folder, f"duplicates_{now}.xlsx")
        duplicates_to_save.to_excel(output_file, index=False)
    print(f"Excel file saved to \n{output_file}")


def delete_duplicates():
    excel_path = input("Enter the path to the file with duplicates:\n")
    # Load the Excel table

    excel_path = excel_path.replace('"', '')

    if not os.path.isfile(excel_path):
        print(f"File {excel_path} not found...")
        return

    df = pd.read_excel(excel_path)

    # Group files by Size and Extension
    grouped = df.groupby(['Size', 'Extension'])

    # Initialize total size counter
    total_deleted_size = 0

    # Iterate through each group
    for (size, extension), group in grouped:
        # Identify the master file
        master_file = group[group['Master'] == True]
        if master_file.empty:
            continue
        master_checksum = master_file['Check_Sum'].values[0]

        # Iterate through files in the group
        for index, row in group.iterrows():
            if not row['Master'] and row['Check_Sum'] == master_checksum:
                file_path = row['File Path']
                try:
                    file_size = os.path.getsize(file_path)
                    os.remove(file_path)
                    total_deleted_size += file_size
                    print(f"Deleted: {file_path} (Size: {file_size} bytes)")
                except OSError as e:
                    print(f"Error deleting {file_path}: {e}")

    total_deleted_size_mb = total_deleted_size / (1024 * 1024)
    print(f"Total size of deleted files: {total_deleted_size_mb:.2f} MB")


if __name__ == "__main__":

    while True:
        action = input("""Select action:
        c - create file of duplicates
        d - delete duplicates based on reviewed file\n""")

        if action.lower() == "c":
            folder_path = input("Enter the folder path to search for duplicates:\n")
            duplicates = find_duplicates(folder_path)
            if duplicates is not None:
                if isinstance(duplicates, pd.DataFrame):
                    if len(duplicates):
                        save_to_excel(duplicates, folder_path)
                    else:
                        print("No duplicates found in the specified folder")
            else:
                print("Something went wrong...")
                time.sleep(5)

        if action.lower() == "d":
            delete_duplicates()

        if action.lower() not in ["c", "d"]:
            print("Invalid action. Please try again\n")

        proceed = input("Select the action:\ne - exit\nanother key - proceed\n")

        if proceed.lower() == "e":
            print("Script is finished")
            time.sleep(5)
            sys.exit(0)
