import pdfplumber
import xlwings as xw

# Open the PDF file
with pdfplumber.open('table.pdf') as pdf:
    # Extract data from the PDF
    data = {}
    for page in pdf.pages:
        for line in page.extract_text().split('\n'):
            row = line.split()  # Assuming the data is space-separated, adjust as needed
            if len(row) >= 2:
                serial_no = row[0]
                name = row[1]
                data[serial_no] = name

# Open the Excel file using xlwings
excel_file = 'data from table.xlsx'
app = xw.App(visible=True)  # This will open Excel if it's not already running

try:
    wb = app.books.open(excel_file)
    sheet = wb.sheets.active  # Select the active sheet, or specify a specific sheet

    # Iterate through the Excel file and insert names into the first column (A)
    for row in sheet.range('A2:A{}'.format(sheet.cells.last_cell.row)):  # Use only the first column
        serial_no = row.value
        if serial_no in data:
            name = data[serial_no]
            row.offset(column_offset=1).value = name  # Offset by 1 column to insert the name

    # Save the modified Excel file
    wb.save()
finally:
    app.quit()
