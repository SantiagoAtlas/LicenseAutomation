import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation
import pdfplumber

# File paths
checklist_file_path = 'C:\\Data\\LicenseAutomation\\System_Checklist_Example.xlsm'
system_list_file_path = 'C:\\Data\\LicenseAutomation\\System_List_Example.xlsx'
pdf_file_path = 'C:\\Data\\LicenseAutomation\\Prüfprotokoll.pdf'

# Headers
headers = [
    'SAP Kommissinsnummer', 'SAP Position', 'Material-Nummer vom Schrank', 'Firma', 'Standort', 'Projekt',
    'Steuerung', 'SW-Paket-Version', 'Komponente', 'Schnittstellen', 'Stationsname', 'IP-Adresse (Kundennetz)',
    'MAC-Adersse 1', 'Lizenznummer (S/N)', 'Lizenz', 'Auslauf-Datum', 'Kommentar'
]

# Function to extract version from PDF
def extract_specific_version(pdf_path, version_name):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and version_name in row[0]:
                        return row[1]
    return None

# Step 1: Extract version
version_name = 'V3 – Board Firmware'
print("🔍 Searching for version in PDF...")
specific_version = extract_specific_version(pdf_file_path, version_name)
if specific_version:
    print(f"✅ Version found for '{version_name}': {specific_version}")
else:
    print(f"⚠️ Version '{version_name}' not found in PDF.")

# Step 2: Load Excel files
print("📂 Opening Excel files...")
checklist_workbook = openpyxl.load_workbook(checklist_file_path, keep_vba=True)
main_info_sheet = checklist_workbook['Main Info']
systemmatrix_sheet = checklist_workbook['Systemmatrix']
system_list_workbook = openpyxl.load_workbook(system_list_file_path)
overview_sheet = system_list_workbook['Overview'] if 'Overview' in system_list_workbook.sheetnames else system_list_workbook.create_sheet('Overview')

# Step 3: Read Main Info values
commission_no = main_info_sheet['B3'].value
end_customer = main_info_sheet['B7'].value
location = main_info_sheet['B8'].value
project_name = main_info_sheet['B5'].value
system = main_info_sheet['B11'].value

print("📋 Retrieved Main Info data:")
print(f"   🔹 Commission No: {commission_no}")
print(f"   🔹 Customer: {end_customer}")
print(f"   🔹 Location: {location}")
print(f"   🔹 Project: {project_name}")
print(f"   🔹 System: {system}")

# Step 4: Handle headers
existing_headers = [cell.value for cell in overview_sheet[1]]
if all(h in existing_headers for h in headers):
    print("✅ Headers already exist in the Overview sheet.")
else:
    overview_sheet.delete_rows(1, 1)
    overview_sheet.append(headers)
    print("🆕 Headers created in the Overview sheet.")
    center_alignment = Alignment(horizontal='center')
    for col in range(1, len(headers) + 1):
        cell = overview_sheet.cell(row=1, column=col)
        cell.alignment = center_alignment
        cell.font = Font(bold=True)

# Step 5: Check and add commission data
print(f"📊 Checking if commission number {commission_no} exists in Overview sheet...")
existing_rows = list(overview_sheet.iter_rows(min_row=2, values_only=True))
existing_positions_for_commission = {
    row[1] for row in existing_rows if row[0] == commission_no
}

added_rows = 0
skipped_rows = 0
start_col_index = column_index_from_string('G')

for col in range(start_col_index, systemmatrix_sheet.max_column + 1):
    sap_position = systemmatrix_sheet.cell(row=4, column=col).value
    if not sap_position:
        continue

    has_x = False
    item_name = order_article_number = None
    for row in range(5, systemmatrix_sheet.max_row + 1):
        val = systemmatrix_sheet.cell(row=row, column=col).value
        if isinstance(val, str) and val.strip().lower() == 'x':
            has_x = True
            item_name = systemmatrix_sheet.cell(row=row, column=2).value
            order_article_number = systemmatrix_sheet.cell(row=row, column=3).value
            break

    if has_x:
        if sap_position in existing_positions_for_commission:
            print(f"⏭️ Position {sap_position} already exists for commission {commission_no}, skipping.")
            skipped_rows += 1
            continue

        combined_value = f"{item_name} / {order_article_number}" if item_name and order_article_number else item_name or order_article_number
        new_row = [
            commission_no, sap_position, combined_value, end_customer, location,
            project_name, system, specific_version, '', '', '', '', '', '', '', '', ''
        ]
        overview_sheet.append(new_row)
        added_rows += 1
        print(f"➕ Added position {sap_position} to commission {commission_no}.")

        # Formatting
        target_row = overview_sheet.max_row
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        right_alignment = Alignment(horizontal='right')
        for col_index in range(1, len(headers) + 1):
            cell = overview_sheet.cell(row=target_row, column=col_index)
            cell.border = thin_border
            cell.alignment = right_alignment

print(f"\n🧾 Summary for Commission {commission_no}:")
print(f"   ✅ New rows added: {added_rows}")
print(f"   ⏭️ Positions skipped (already existed): {skipped_rows}")

# Step 6: Add data validation
print("🔧 Adding data validation for 'Interfaces' column...")
data_validation = DataValidation(type="list", formula1='"MQTT,OPC UA,IBM MQ"', showDropDown=True)
overview_sheet.add_data_validation(data_validation)
for row in range(2, overview_sheet.max_row + 1):
    cell = overview_sheet[f'I{row}']
    data_validation.add(cell)
    cell.border = thin_border
    cell.alignment = right_alignment

# Step 7: Adjust column widths
print("📐 Adjusting column widths...")
for col in range(1, len(headers) + 1):
    max_length = max((len(str(cell.value)) if cell.value else 0) for cell in overview_sheet[get_column_letter(col)])
    overview_sheet.column_dimensions[get_column_letter(col)].width = max_length + 2

overview_sheet.auto_filter.ref = overview_sheet.dimensions

# Step 8: Save file
try:
    system_list_workbook.save(system_list_file_path)
    print("💾 File saved successfully. Process completed ✅")
except PermissionError:
    print("❌ Error: Could not save file. Please close the Excel file if it's open.")
