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

# Find the next empty row in the "Comision Number" column
comision_column = system_list_sheet['B']
next_empty_row = len(comision_column) + 1

# Write the Commission No. to the next empty row in the "Comision Number" column
system_list_sheet[f'B{next_empty_row}'] = commission_no

# === NUEVA LÓGICA CON HOJA 'Systemmatrix' ===
# Imprimir todos los valores desde G4 hasta el final de la fila 4 en Systemmatrix
start_col_index = column_index_from_string('G')  # G = 7
row_index = 4

print("Valores encontrados en la hoja 'Systemmatrix', fila 4 desde la columna G en adelante:")

for col in range(start_col_index, systemmatrix_sheet.max_column + 1):
    cell = systemmatrix_sheet.cell(row=row_index, column=col)
    value = cell.value
    print(f"- {value}")

# Guardar cambios en el archivo de destino
system_list_workbook.save(system_list_file_path)

# Confirmación final
print(f"\nCommission No. {commission_no} has been appended to the 'System List Example.xlsx' file.")
