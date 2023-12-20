import csv
from openpyxl import Workbook
import os
import sys
from config.confidential_info import operating_file_path_instance

# Define replacements dictionary
replacements = {"NT": 0, "ND": 0, "PASS": 0}

# Downloads folder path
downloads_folder_path = os.path.join(os.path.expanduser("~"), "Downloads")

# Find the downloaded file
matching_files = [f for f in os.listdir(
    downloads_folder_path) if f.startswith("NA") and f.endswith(".csv")]

if not matching_files:
    print("No matching CSV files found in Downloads folder.")
    sys.exit()
else:
    # Assume the first downloaded file is the relevant one
    downloaded_file_path = os.path.join(
        downloads_folder_path, matching_files[0])

    # Open the CSV file for processing
    with open(downloaded_file_path, 'r') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)  # skip header row
        data = list(reader)

    # Replace specific values in data
    for row_index, row in enumerate(data):
        for col_index, cell_value in enumerate(row):
            if cell_value in replacements:
                data[row_index][col_index] = replacements[cell_value]

    # Create new workbook and sheet
    wb1 = Workbook()
    ws1 = wb1.active

    # Copy processed data to new workbook
    data_row = 1
    for row in data:
        data_col = 1
        for cell_value in row:
            ws1.cell(row=data_row, column=data_col).value = cell_value
            data_col += 1
        data_row += 1

    # Generate the processed file path
    processed_file_path = os.path.splitext(downloaded_file_path)[
        0] + "_processed.xlsx"

    # Save new workbook with dynamically generated filename
    wb1.save(processed_file_path)

    print(
        f"Data processed successfully. New workbook saved as: {processed_file_path}")

os.system(operating_file_path_instance.add_to_local)
print("Chained script add to local execution completed.")
