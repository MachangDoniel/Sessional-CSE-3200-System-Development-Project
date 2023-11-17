import pandas as pd
import tabula
from openpyxl import Workbook

def extract_table_from_pdf(pdf_path):
    # Use tabula to extract tables from the PDF
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    return tables

def create_excel_file(name, data):
    # Create a new Excel workbook and add the data to a sheet
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Data'

    for row in data.itertuples(index=False):
        sheet.append(row)

    # Save the Excel file with the given name
    workbook.save(f'{name}.xlsx')

def main(pdf_path):
    # Extract tables from the PDF
    tables = extract_table_from_pdf(pdf_path)

    # Assuming the first table contains the names for Excel files
    first_table = tables[0]
    name_column = first_table.iloc[:, 0]  # Assuming the names are in the first column

    # Iterate through the remaining tables and create Excel files based on names
    for i in range(1, len(tables)):
        current_table = tables[i]
        current_table_name = name_column[i - 1]

        # Create an Excel file for each name and add data from the corresponding table
        create_excel_file(str(current_table_name), current_table)

if __name__ == "__main__":
    pdf_path = "table.pdf"  # Replace with the path to your PDF file
    main(pdf_path)
