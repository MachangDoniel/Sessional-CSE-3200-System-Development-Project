import re
from docx import Document

def extract_department_line(doc):
    pattern = r'(?:Department of|Department Of)(.*)'

    for paragraph in doc.paragraphs:
        match = re.search(pattern, paragraph.text)
        if match:
            department_text = match.group(1).strip()  # Extract text after the department pattern
            return department_text

    return None

# Path to your document
doc_path = r'E:\New\Demo.docx'

try:
    doc = Document(doc_path)
    department_line = extract_department_line(doc)
    
    if department_line:
        print("Department Line:", department_line)
    else:
        print("No 'Department of' line found.")
except FileNotFoundError:
    print(f"Error: The file '{doc_path}' was not found.")
except Exception as e:
    print(f"Error: {e}")
