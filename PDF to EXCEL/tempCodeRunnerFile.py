from PyPDF2 import PdfReader

reader = PdfReader("sample2.pdf")
number_of_pages = len(reader.pages)
page = reader.pages[0]
text=page.extract_text()
print(page.extract_text())