import os
import pandas as pd
from datetime import datetime
import hashlib


def file_checksum(file_path):
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
    except IOError:
        print(f"Не удалось открыть {file_path}")
        return None
    return hash_md5.hexdigest()


def find_duplicates(folder_path):
    data = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_size = os.path.getsize(file_path)
            file_ext = os.path.splitext(file)[1]
            data.append([file_size, file_ext, file_path, file])

    df = pd.DataFrame(data, columns=['Size', 'Extension', 'File Path', 'File Name'])

    duplicates = df[df.duplicated(subset=['Size', 'Extension'], keep=False)].copy()

    # Mark the file with the shortest name as 'master'
    duplicates['Master'] = duplicates.groupby(['Size', 'Extension'])['File Name'].transform(
        lambda x: x == x.loc[x.str.len().idxmin()])

    duplicates['Check_Sum'] = duplicates['File Path'].apply(file_checksum)

    print(duplicates.Check_Sum)

    return duplicates


def save_to_excel(duplicates, output_folder):
    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_file = os.path.join(output_folder, f"duplicates_{now}.xlsx")
    duplicates.to_excel(output_file, index=False)
    print(f"Excel file saved to {output_file}")


if __name__ == "__main__":
    folder_path = input("Enter the folder path to search for duplicates: ")
    duplicates = find_duplicates(folder_path)
    save_to_excel(duplicates, folder_path)
