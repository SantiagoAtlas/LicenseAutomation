import openpyxl
from openpyxl.utils import column_index_from_string

# Paths to the Excel files
checklist_file_path = 'C:\\Data\\LicenseAutomation\\System_Checklist_Example.xlsm'
system_list_file_path = 'C:\\Data\\LicenseAutomation\\System_List_Example.xlsx'

# Load the checklist Excel file
checklist_workbook = openpyxl.load_workbook(checklist_file_path, keep_vba=True)
main_info_sheet = checklist_workbook['Main Info']
systemmatrix_sheet = checklist_workbook['Systemmatrix']

# Extract the Commission No., End Customer, Location, Project Name, and System from 'Main Info'
commission_no = main_info_sheet['B3'].value
end_customer = main_info_sheet['B7'].value
location = main_info_sheet['B8'].value
project_name = main_info_sheet['B5'].value
system = main_info_sheet['B11'].value

# Load the system list Excel file
system_list_workbook = openpyxl.load_workbook(system_list_file_path)
system_list_sheet = system_list_workbook['Übersicht der Systeme']

# === NEW LOGIC WITH 'Systemmatrix' SHEET ===
start_col_index = column_index_from_string('G')  # Start from column G (7)
row_sap_header = 4  # SAP positions are in row 4
first_data_row = 5  # First row below SAP header to check for 'x'

# Find the next empty row in column B
b_column = system_list_sheet['B']
next_empty_row_b = len(b_column) + 1

# Loop through columns G to the last used column
for col in range(start_col_index, systemmatrix_sheet.max_column + 1):
    sap_position = systemmatrix_sheet.cell(row=row_sap_header, column=col).value

    # Skip if no SAP position is defined
    if not sap_position:
        continue

    # Check if there's any 'x' (case-insensitive) below this column
    has_x = False
    item_name = None
    order_article_number = None
    for row in range(first_data_row, systemmatrix_sheet.max_row + 1):
        cell_value = systemmatrix_sheet.cell(row=row, column=col).value
        if isinstance(cell_value, str) and cell_value.strip().lower() == 'x':
            has_x = True
            item_name = systemmatrix_sheet.cell(row=row, column=2).value  # Assuming item name is in column B (2)
            order_article_number = systemmatrix_sheet.cell(row=row, column=3).value  # Assuming order article number is in column C (3)
            break

    # If there's at least one 'x', write to the target sheet
    if has_x:
        combined_value = f"{item_name} / {order_article_number}" if item_name and order_article_number else item_name or order_article_number
        system_list_sheet[f'B{next_empty_row_b}'] = commission_no  # Column B
        system_list_sheet[f'C{next_empty_row_b}'] = sap_position   # Column C
        system_list_sheet[f'D{next_empty_row_b}'] = combined_value # Column D
        system_list_sheet[f'E{next_empty_row_b}'] = end_customer   # Column E
        system_list_sheet[f'H{next_empty_row_b}'] = location       # Column H
        system_list_sheet[f'I{next_empty_row_b}'] = project_name   # Column I
        system_list_sheet[f'J{next_empty_row_b}'] = system         # Column J
        next_empty_row_b += 1

# Save changes
system_list_workbook.save(system_list_file_path)

# Final confirmation
print("\nOnly SAP positions with at least one 'x' were copied to the system list, including End Customer, Item Name, Order Article Number, Location, Project Name, and System.")
