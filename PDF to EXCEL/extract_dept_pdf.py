import re
import pdfplumber

def extract_department_line(pdf_path):
    pattern = r'(?:Department of|Department Of)(.*)'

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            matches = re.findall(pattern, text)
            if matches:
                department_text = matches[0].strip()  # Extract text after the department pattern
                return department_text

    return None

# Path to your PDF document
pdf_path = r'E:\New\Demo.pdf'

try:
    department_line = extract_department_line(pdf_path)
    
    if department_line:
        print("Department Line:", department_line)
    else:
        print("No 'Department of' line found.")
except FileNotFoundError:
    print(f"Error: The file '{pdf_path}' was not found.")
except Exception as e:
    print(f"Error: {e}")
