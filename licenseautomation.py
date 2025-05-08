import openpyxl
from openpyxl.utils import column_index_from_string

# Paths to the Excel files
checklist_file_path = 'C:\\Data\\LicenseAutomation\\System_Checklist_Example.xlsm'
system_list_file_path = 'C:\\Data\\LicenseAutomation\\System_List_Example.xlsx'

# Load the checklist Excel file
checklist_workbook = openpyxl.load_workbook(checklist_file_path, keep_vba=True)
main_info_sheet = checklist_workbook['Main Info']
systemmatrix_sheet = checklist_workbook['Systemmatrix']

# Extract the Commission No. from 'Main Info'
commission_no = main_info_sheet['B3'].value

# Load the system list Excel file
system_list_workbook = openpyxl.load_workbook(system_list_file_path)
system_list_sheet = system_list_workbook['Übersicht der Systeme']

# === NEW LOGIC WITH 'Systemmatrix' SHEET ===
# Extract all values from G4 onwards in Systemmatrix
start_col_index = column_index_from_string('G')  # G = 7
row_index = 4

# Collect the SAP positions from row 4, starting from column G
sap_positions = []
for col in range(start_col_index, systemmatrix_sheet.max_column + 1):
    cell = systemmatrix_sheet.cell(row=row_index, column=col)
    value = cell.value
    if value:  # Only add non-empty values
        sap_positions.append(value)

# Write the 'commission number' in column B and 'sap positions' in column C
# Find the next empty row in column B of 'Übersicht der Systeme'
b_column = system_list_sheet['B']
next_empty_row_b = len(b_column) + 1

# Write the 'commission number' and 'sap position' in the corresponding rows
for sap_position in sap_positions:
    system_list_sheet[f'B{next_empty_row_b}'] = commission_no  # Write Commission No. in column B
    system_list_sheet[f'C{next_empty_row_b}'] = sap_position   # Write SAP Position in column C
    next_empty_row_b += 1

# Save changes to the destination file
system_list_workbook.save(system_list_file_path)

# Final confirmation
print(f"\nThe Commission No. has been added to column B and the SAP Position to column C of 'System List Example.xlsx'.")
