import openpyxl
import requests
from openpyxl.styles import PatternFill

# Function to check if a Google Drive link is accessible
def is_google_drive_accessible(url):
    try:
        response = requests.head(url, allow_redirects=True)
        # Check if the response status code indicates accessibility
        return response.status_code in [200, 403, 404]
    except requests.RequestException:
        return False

# Function to process the Excel file
def check_lke_excel(file_path, output_file_path):
    # Load the workbook and select the "jawaban" sheet
    wb = openpyxl.load_workbook(file_path)
    sheet = wb["jawaban"]
    
    # Prepare to write the output
    output_wb = openpyxl.Workbook()
    output_sheet = output_wb.active
    output_sheet.title = "jawaban"
    
    # Iterate through the rows in the sheet
    for row in range(2, sheet.max_row + 1):  # Start from the second row
        # Check column K
        k_cell = sheet.cell(row=row, column=11)  # Column K
        m_cell = sheet.cell(row=row, column=13)  # Column M
        r_cell = sheet.cell(row=row, column=18)  # Column R

        if k_cell.value is None:
            k_cell.fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
            r_cell.value = "keterangan bukti dukung"
        else:
            k_cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        # Check column M for Google Drive links
        if m_cell.value:
            if is_google_drive_accessible(m_cell.value):
                r_cell.value = "keterangan pengisian jawaban"
            else:
                r_cell.value = "Link not accessible"
        else:
            r_cell.value = "Cell is empty"

    # Save the output Excel file
    output_wb.save(output_file_path)  

# Example usage
# check_lke_excel('input_file.xlsx', 'output_file.xlsx')
