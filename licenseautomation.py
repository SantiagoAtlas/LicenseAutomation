import openpyxl

# Paths to the Excel files
checklist_file_path = 'C:\\Data\\LicenseAutomation\\System_Checklist_Example.xlsm'
system_list_file_path = 'C:\\Data\\LicenseAutomation\\System_List_Example.xlsx'

# Load the checklist Excel file
checklist_workbook = openpyxl.load_workbook(checklist_file_path, keep_vba=True)
checklist_sheet = checklist_workbook['Main Info']

# Extract the Commission No.
commission_no = checklist_sheet['B3'].value

# Load the system list Excel file
system_list_workbook = openpyxl.load_workbook(system_list_file_path)
system_list_sheet = system_list_workbook['Übersicht der Systeme']

# Find the next empty row in the "Comision Number" column
comision_column = system_list_sheet['B']  
next_empty_row = len(comision_column) + 1

# Write the Commission No. to the next empty row in the "Comision Number" column
system_list_sheet[f'B{next_empty_row}'] = commission_no

# Save the updated system list Excel file
system_list_workbook.save(system_list_file_path)

# Print confirmation message
print(f"Commission No. {commission_no} has been appended to the 'Comision Number' column in the 'System List Example.xlsx' file.")
