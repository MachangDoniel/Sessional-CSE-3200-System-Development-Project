from openpyxl import Workbook

def write_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active

    for table in data:
        for row in table:
            ws.append(row)

    wb.save(output_file)
