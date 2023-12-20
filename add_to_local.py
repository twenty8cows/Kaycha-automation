import os
from openpyxl import load_workbook
from config.confidential_info import operating_file_path_instance

# # Path where the Excel file is located
operating_file_path = operating_file_path_instance.operating_file_path
excel_filename = operating_file_path_instance.excel_filename  # The provided Excel file
full_excel_path = os.path.join(operating_file_path, excel_filename)

# Load the existing Excel workbook
workbook = load_workbook(full_excel_path, keep_vba=True)
worksheet = workbook["Kaycha"]  # Access the 'Kaycha' worksheet

# Find the first empty cell in column A
first_empty_row = 1
for row in worksheet.iter_rows(min_col=1, max_col=1):
    if row[0].value is None:
        break
    first_empty_row += 1

# Find the file in the downloads folder that ends with _processed.xlsx
downloads_folder_path = os.path.expanduser('~/Downloads')
matching_files = [f for f in os.listdir(
    downloads_folder_path) if f.endswith('_processed.xlsx')]

if not matching_files:
    print("No processed file found")
    exit()  # Exit the script if no file is found
else:
    processed_file_path = os.path.join(
        downloads_folder_path, matching_files[0])
    print(f"Found processed file: {processed_file_path}")

# Load the processed workbook
processed_workbook = load_workbook(processed_file_path)
# Assuming data is in the active sheet
processed_sheet = processed_workbook.active

# Append the data from the processed sheet to the destination worksheet
# Skip header if present
for row in processed_sheet.iter_rows(min_row=2, values_only=True):
    for col_index, cell_value in enumerate(row, start=1):
        worksheet.cell(row=first_empty_row, column=col_index, value=cell_value)
    first_empty_row += 1

# Save the updated workbook
workbook.save(full_excel_path)
print(f"Data appended successfully to {full_excel_path}")
