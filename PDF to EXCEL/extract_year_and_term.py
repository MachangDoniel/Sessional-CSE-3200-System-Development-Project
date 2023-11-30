import re
from docx import Document

def extract_year_term_before_table(doc):
    pattern_bills = r'Bills - (\w+).*?year (\w+)'
    pattern_exam = r'Examination- (\w+)'

    found_bills = found_exam = False
    word_after_bills = word_after_year = word_after_exam = None

    for paragraph in doc.paragraphs:
        match_bills = re.search(pattern_bills, paragraph.text)
        match_exam = re.search(pattern_exam, paragraph.text)
        
        if match_bills:
            found_bills = True
            word_after_bills = match_bills.group(1)
            word_after_year = match_bills.group(2)

        if match_exam:
            found_exam = True
            word_after_exam = match_exam.group(1)

        if found_bills and found_exam:
            return word_after_bills, word_after_year, word_after_exam

    return word_after_bills, word_after_year, word_after_exam

doc_path = r'E:\New\Demo.docx'

try:
    doc = Document(doc_path)
    result_bills, result_year, result_exam = extract_year_term_before_table(doc)
    
    if result_bills and result_year and result_exam:
        print("Word after 'bills -':", result_bills)
        print("Word after 'year':", result_year)
        print("Word after 'Examination-':", result_exam)
    elif result_bills and result_year:
        print("Word after 'bills -':", result_bills)
        print("Word after 'year':", result_year)
    elif result_exam:
        print("Word after 'Examination-':", result_exam)
    else:
        print("No match found before the table.")
except FileNotFoundError:
    print(f"Error: The file '{doc_path}' was not found.")
except Exception as e:
    print(f"Error: {e}")
