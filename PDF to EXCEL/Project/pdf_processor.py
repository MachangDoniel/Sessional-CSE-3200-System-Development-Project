import pdfplumber

def extract_table_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        tables = []
        for page in pdf.pages:
            table = page.extract_table()
            tables.append(table)
        return tables
