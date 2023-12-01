from docx import Document

def extract_tables_with_titles(file_path):
    doc = Document(file_path)
    tables = doc.tables
    titles = {}

    for table in tables:
        if table.rows:
            first_line = table.rows[0].cells[0].text.strip()
            table_content = '\n'.join([cell.text.strip() for row in table.rows for cell in row.cells])
            
            # Get the text before the table as the title
            previous_paragraph = table._tbl.getprevious()
            if previous_paragraph is not None:
                title = previous_paragraph.text.strip()
                print("Title: ",title)
                titles[title] = table_content
        print("Table: ",table)
    return titles

# Replace 'your_file.docx' with the path to your Word file
file_path = 'Demo.docx'
tables_with_titles = extract_tables_with_titles(file_path)
count=0
for title, table_content in tables_with_titles.items():
    print(f"Title {count} :", title)
    count+=1
    # print("Table:")
    # print(table_content)
    print()