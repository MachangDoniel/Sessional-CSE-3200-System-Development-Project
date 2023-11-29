# PDF to EXCEL
The project is about to build a system that read a pdf file, extract data from it and write it in excel file.
## Installation
1. Download [VS code](https://code.visualstudio.com/)
## Environment SetUp
We need some [Python](https://www.python.org/downloads/) libraries.
To work with Python in [VS code](https://code.visualstudio.com/), we should install the Python extension.

a. Click on the "Extensions" icon in the sidebar on the left (or use the keyboard shortcut Ctrl+Shift+X on Windows/Linux or Cmd+Shift+X on macOS).

b. In the search bar at the top of the Extensions pane, type "Python."

c. Find the "Python" extension by Microsoft and click the "Install" button.
1. Install [pypdf2](https://pypi.org/project/PyPDF2/) & [pdfplumber](https://pypi.org/project/pdfplumber/) for PDF extraction.
2. Install [openpyxl](https://openpyxl.readthedocs.io/en/stable/index.html) & [pandas](https://pandas.pydata.org/) for Excel manipulation.
3. Install [matplotlib](https://matplotlib.org/) for Visualization

You can install these features via pip, write the command in Teminal
```bash
pip install pypdf2
pip install pdfplumber
pip install openpyxl
pip install pandas
pip install matplotlib
```
To update these features, use the following command
```bash
pip install --upgrade pip
```
## Working With pypdf2
### Extract data from pdf file
The following code will read a page(0th) of simple text from a source file table with text.pdf(replace it with the name of your source file)
```bash
from PyPDF2 import PdfReader

reader = PdfReader("table with text.pdf")
page = reader.pages[0]
print(page.extract_text())
```
![Alt text](image-1.png)

we can limit the text orientation by using the following code instead of "print(page.extract_text())"
```bash
# extract only text oriented up
print(page.extract_text(0))

# extract text oriented up and turned left
print(page.extract_text((0, 90)))
```

## Working With pdfplumber

```bash
import pdfplumber

with pdfplumber.open("text.pdf") as pdf:
    first_page = pdf.pages[0]
    print(first_page.chars[0])
```
![Alt text](image-5.png)

### Extract text from the PDF:

To extract all the text from the PDF, we can iterate through each page and extract the text as follows:

```bash
import pdfplumber

with pdfplumber.open("text.pdf") as pdf:
    text = ''
    for page in pdf.pages:
        text += page.extract_text()
        print(text)
```
![Alt text](image-4.png)
We can also extract text from a specific page by using pdf.pages[i], where i is the page number (0-based index).
```bash
import pdfplumber

with pdfplumber.open("text.pdf") as pdf:
    text=''
    text += pdf.pages[0].extract_text()
    print(text)
```
![Alt text](image-6.png)

### Extract table data:

If the PDF contains tables, you can extract table data as well. pdfplumber allows you to extract tables from PDFs as Pandas DataFrames, making it easy to work with tabular data.

**Note:** If the pdf file contains table within text, the following code can simply ignore it extracting only the table data.
```bash
import pdfplumber
import pandas as pd

# Open the PDF file
with pdfplumber.open('table.pdf') as pdf:
    # Extract a table from a specific page
    page = pdf.pages[0]  # You can specify the page number
    table = page.extract_table()

# Convert the extracted table data to a Pandas DataFrame
df = pd.DataFrame(table)

# Print the DataFrame
print(df)
```
![Alt text](image-7.png)

### Save the table as csv file
```bash
df.to_csv('table_data.csv', index=False)
```
#### csv file output
```bash
Serial No,Students Name
1.,Aditi Chakma
2.,Anik Ekka
3.,Arnab Talukdar
4.,Darpan Chakma
```
![Alt text](image-13.png)
### Save the table as excel file
```bash
df.to_excel('table_data.xlsx', index=False)
```
![Alt text](image-8.png)
![Alt text](image-9.png)
### Save the table as a 2D matrix 
If you want to save the table data from a PDF as a 2D matrix (i.e., a list of lists), you can do so by simply converting the extracted table into a 2D list

```bash
import pdfplumber

# Open the PDF file
with pdfplumber.open('table.pdf') as pdf:
    # Extract a table from a specific page
    page = pdf.pages[0]  # You can specify the page number
    table = page.extract_table()

# Convert the extracted table data into a 2D matrix (list of lists)
matrix = [list(row) for row in table]

# Print the matrix (for verification)
for row in matrix:
    print(row)
```
![Alt text](image.png)
### Save the 2D matrix as csv file
```bash
import pdfplumber
import csv

# Open the PDF file
with pdfplumber.open('table.pdf') as pdf:
    # Extract a table from a specific page
    page = pdf.pages[0]  # You can specify the page number
    table = page.extract_table()

# Convert the extracted table data into a 2D matrix (list of lists)
matrix = [list(row) for row in table]

# Print the matrix (for verification)
for row in matrix:
    print(row)

with open('matrix_data.csv', 'w', newline='') as csvfile:
    csvwriter = csv.writer(csvfile)
    for row in matrix:
        csvwriter.writerow(row)

```
### Save the 2D matrix as excel file
```bash
import pdfplumber
import pandas as pd

# Open the PDF file
with pdfplumber.open('table.pdf') as pdf:
    # Extract a table from a specific page
    page = pdf.pages[0]  # You can specify the page number
    table = page.extract_table()

# Convert the extracted table data into a 2D matrix (list of lists)
matrix = [list(row) for row in table]

# Print the matrix (for verification)
for row in matrix:
    print(row)

df = pd.DataFrame(matrix)

# Save the DataFrame to an Excel file
df.to_excel('matrix_data.xlsx', index=False, header=False)
```
![Alt text](image-10.png)
![Alt text](image-12.png)
#### Install Rainbow CSV extension to watch over the csv file directly
![Alt text](image-11.png)

### Write extracting data from pdf to excel
Here, Initially I have a pdf file containing serial no and name of some students and an excel file containing only serial no, i want to enter the name into it.

![Alt text](image-14.png)
![Alt text](image-15.png)

```bash
import pdfplumber
from openpyxl import load_workbook

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

# Load the Excel file
excel_file = 'data from table.xlsx'
wb = load_workbook(excel_file)
sheet = wb.active  # Select the active sheet, you can choose a specific sheet if needed

# Iterate through the Excel file and insert names
for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=2):
    serial_no = row[0].value
    if serial_no in data:
        name = data[serial_no]
        row[1].value = name

# Save the modified Excel file
wb.save(excel_file)
```
![Alt text](image-16.png)
![Alt text](image-17.png)
**Note:** Before running the program, the excel file needs to be closed.
When an Excel file is open in an application, it is often locked for editing by other processes, including external scripts. That's why we may experience the issue where our Python script can't save the Excel file when it is open in Excel.
or simply we can use xlwings library or pywin32 to autmate Excel from Python script. Those libraries help us to manipulate the Excel file including writing data, even when they are open in Excel.
```bash
pip install xlwings
```
```bash
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

```
## Lets Jump into the Project
```bash
import pdfplumber

# Open the PDF file
with pdfplumber.open('Exam Bill Demo.pdf') as pdf:
    table = None

    # Iterate through all pages
    for page in pdf.pages:
        # Attempt to extract a table from the page
        extracted_table = page.extract_table()

        # Check if a table was successfully extracted
        if extracted_table:
            table = extracted_table
            break  # Exit the loop if a table is found on any page

    if table:
        # Iterate over the rows and columns of the table
        for row in table:
            for cell in row:
            # cell=row[2]
                print(cell, end='\t')  # Printing the cell value, separate columns with tabs
            print()  # Start a new line for the next row
```
## Lets print all the tables from pdf file
```bash
import pdfplumber

# Open the PDF file
with pdfplumber.open('Exam Bill Demo.pdf') as pdf:
    tables = []  # Initialize a list to store tables

    # Iterate through all pages
    for page in pdf.pages:
        # Attempt to extract a table from the page
        extracted_table = page.extract_table()

        # Check if a table was successfully extracted
        if extracted_table:
            tables.append(extracted_table)  # Append the table to the list

    # Process the accumulated tables (if any)
    for table in tables:
        # Iterate over the rows and columns of each table
        for row in table:
            for cell in row:
                print(cell, end='\t')  # Printing the cell value, separate columns with tabs
            print()  # Start a new line for the next row
        print()
    print()
```
we need to identify each table seperately, and accoding to its title, we need to extract data.
```bash
import pdfplumber

# Open the PDF file
with pdfplumber.open('Exam Bill Demo.pdf') as pdf:
    tables = []  # Initialize a list to store tables

    # Iterate through all pages
    for page in pdf.pages:
        # Attempt to extract a table from the page
        extracted_table = page.extract_table()

        # Check if a table was successfully extracted
        if extracted_table:
            tables.append(extracted_table)  # Append the table to the list

    # Process the accumulated tables (if any)
    table_no=0
    for table in tables:
        # Iterate over the rows and columns of each table
        table_no+=1
        print("Table no: "+str(table_no))
        for row in table:
            for cell in row:
                print(cell, end='\t')  # Printing the cell value, separate columns with tabs
            print()  # Start a new line for the next row
        print()
    print()
```

## latest
```bash
import pdfplumber

# Open the PDF file
with pdfplumber.open('Exam Bill Demo.pdf') as pdf:
    preceding_line = None  # Initialize variable to store the preceding line
    table_no = 0  # Initialize the table number

    # Iterate through all pages
    for page in pdf.pages:
        # Extract text content from the page
        page_text = page.extract_text()
        
        # Check if the page contains a table
        if page.extract_table():
            # Increment the table number
            table_no += 1
            # Store the text immediately preceding the table
            preceding_line = page_text

        # If we have a stored preceding line, process it and the table
        if preceding_line:
            print()
            print()
            print("Table no:", table_no)
            print()
            print()
            print("Preceding Line:", preceding_line)
            print()
            print()
            print()
            print("Table Start:")
            print()
            print()
            table = page.extract_table()
            if table:
                for row in table:
                    for cell in row:
                        print(cell, end='\t')
                    print()
            print()
            preceding_line = None
```

# Environment to make Specific PDF to Many excel files
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import os

def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text_content = ""
        tables_with_titles = []
        for page in pdf.pages:
            text_content += page.extract_text()
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    # Extracting the title before the table
                    table_start = table[0][1] if table else 0
                    text_before_table = page.extract_text()
                    lines = text_before_table.split('\n')
                    title = ""
                    for line in reversed(lines):
                        if line.strip():
                            title = line
                            break
                    tables_with_titles.append({"Title": title, "Table": table})
    return text_content, tables_with_titles


def select_pdf():
    global pdf_file
    pdf_file = filedialog.askopenfilename()
    if pdf_file:
        pdf_label.config(text=f"Selected PDF: {pdf_file}")
        pdf_label.pack()


def generate_excel():
    global pdf_file
    if pdf_file:
        text_content, tables_with_titles = extract_data_from_pdf(pdf_file)
        
        output_dir = filedialog.askdirectory()
        print(output_dir)  # Debug line to check the selected directory
        
        if output_dir:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)  # Create the directory if it doesn't exist
            if tables_with_titles:
                for idx, data in enumerate(tables_with_titles):
                    title = data["Title"]
                    table = data["Table"]
                    print(f"Title: {title}")  # Print extracted title to console
                    df = pd.DataFrame(table)
                    df.ffill(axis=0, inplace=True)  # Fill empty cells with previous values in the same column
                    excel_path = f"{output_dir}/table_{idx}.xlsx"  # Using table number for filename
                    try:
                        df.to_excel(excel_path, index=False)
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to create Excel file at {excel_path}: {e}")
                        print(f"Failed to create Excel file at {excel_path}: {e}")  # Detailed error message in console
                        break
                else:
                    messagebox.showinfo("Excel Created Successfully", f"Excel(s) created successfully in {output_dir}!")
            else:
                messagebox.showwarning("No Tables Found", "No tables were detected in the PDF.")
        else:
            messagebox.showwarning("Debug", "No output directory selected.")




def refresh_app():
    root.destroy()
    main()


def main():
    global root, pdf_label

    root = tk.Tk()
    root.title("PDF to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_button = ttk.Button(main_frame, text="Select PDF", command=select_pdf, style="TButton")
    select_button.pack(pady=10)

    pdf_label = tk.Label(main_frame, text="Selected PDF: ", bg="#f0f0f0")
    pdf_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Excel", command=generate_excel, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
```
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import os

def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text_content = ""
        tables_with_titles = []
        for page in pdf.pages:
            text_content += page.extract_text()
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    # Extracting the title before the table
                    table_start = table[0][1] if table else 0
                    text_before_table = page.extract_text()
                    lines = text_before_table.split('\n')
                    title = ""
                    for line in reversed(lines):
                        if line.strip():
                            title = line
                            break
                    tables_with_titles.append({"Title": title, "Table": table})
    return text_content, tables_with_titles


def select_pdf():
    global pdf_file
    pdf_file = filedialog.askopenfilename()
    if pdf_file:
        pdf_label.config(text=f"Selected PDF: {pdf_file}")
        pdf_label.pack()


def generate_excel():
    global pdf_file
    if pdf_file:
        text_content, tables_with_titles = extract_data_from_pdf(pdf_file)
        
        output_dir = filedialog.askdirectory()
        print(output_dir)  # Debug line to check the selected directory
        
        if output_dir:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)  # Create the directory if it doesn't exist
            if tables_with_titles:
                for idx, data in enumerate(tables_with_titles):
                    title = data["Title"]
                    table = data["Table"]
                    df = pd.DataFrame(table)
                    df.ffill(axis=0, inplace=True)  # Fill empty cells with previous values in the same column
                    excel_path = f"{output_dir}/table_{idx}.xlsx"  # Using table number for filename
                    try:
                        df.to_excel(excel_path, index=False)
                        print(f"table_{idx}.xlsx created")  # Print the Excel file created
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to create Excel file at {excel_path}: {e}")
                        print(f"Failed to create Excel file at {excel_path}: {e}")  # Detailed error message in console
                        break
                
                messagebox.showinfo("Excel Created Successfully", f"Excel(s) created successfully in {output_dir}!")
                clear_labels()  # Reset the labels after generating Excel files
            else:
                messagebox.showwarning("No Tables Found", "No tables were detected in the PDF.")
        else:
            messagebox.showwarning("Debug", "No output directory selected.")



def clear_labels():
    pdf_label.config(text="Selected PDF: ")


def refresh_app():
    root.destroy()
    main()


def main():
    global root, pdf_label

    root = tk.Tk()
    root.title("PDF to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_button = ttk.Button(main_frame, text="Select PDF", command=select_pdf, style="TButton")
    select_button.pack(pady=10)

    pdf_label = tk.Label(main_frame, text="Selected PDF: ", bg="#f0f0f0")
    pdf_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Excel", command=generate_excel, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
```
# Creating The excel file for each Teacher
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import os
from shutil import copyfile


def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        tables_with_titles = []
        page = pdf.pages[0]  # Considering only the first page
        tables = page.extract_tables()
        if tables:
            first_table = tables[0]
            tables_with_titles.append({"Table": first_table})
    return tables_with_titles



def select_pdf():
    global pdf_file
    pdf_file = filedialog.askopenfilename()
    if pdf_file:
        pdf_label.config(text=f"Selected PDF: {pdf_file}")
        pdf_label.pack()


def select_sample_excel():
    global sample_file
    sample_file = filedialog.askopenfilename()
    if sample_file:
        sample_label.config(text=f"Selected Sample Excel: {sample_file}")
        sample_label.pack()


def copy_and_edit_excel_files(tables_with_titles):
    if not tables_with_titles:
        messagebox.showwarning("No Tables Found", "No tables were detected in the PDF.")
        return

    output_dir = filedialog.askdirectory()
    if not output_dir:
        messagebox.showwarning("Debug", "No output directory selected.")
        return

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for idx, data in enumerate(tables_with_titles):
        table = data["Table"]
        if table:
            # Considering the first table only for file creation
            df = pd.DataFrame(table)
            # Skipping the header row and starting from the second row
            for i in range(1, len(df)):
                row_values = df.iloc[i].tolist()
                file_name = f"{row_values[0]}_{row_values[1]}.xlsx"
                try:
                    # Copy and edit the sample file
                    new_file_path = os.path.join(output_dir, file_name)
                    copyfile(sample_file, new_file_path)
                    print(f"{file_name} created")  # Print the file name created
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to create Excel file at {file_name}: {e}")
                    print(f"Failed to create Excel file at {file_name}: {e}")  # Detailed error message in console

    messagebox.showinfo("Excel Files Created", f"Excel files created successfully in {output_dir}!")
    clear_labels()  # Reset the labels after generating Excel files


def generate_excel():
    global pdf_file
    if pdf_file:
        tables_with_titles = extract_data_from_pdf(pdf_file)
        copy_and_edit_excel_files(tables_with_titles)


def clear_labels():
    pdf_label.config(text="Selected PDF: ")
    sample_label.config(text="Selected Sample Excel: ")


def refresh_app():
    root.destroy()
    main()


def main():
    global root, pdf_label, sample_label

    root = tk.Tk()
    root.title("PDF to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_pdf_button = ttk.Button(main_frame, text="Select PDF", command=select_pdf, style="TButton")
    select_pdf_button.pack(pady=10)

    pdf_label = tk.Label(main_frame, text="Selected PDF: ", bg="#f0f0f0")
    pdf_label.pack()

    select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=select_sample_excel, style="TButton")
    select_sample_button.pack(pady=10)

    sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
    sample_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Excel", command=generate_excel, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
```
# Make a temp118121 folder and place all the table into seperate excel files & make excel file from the first table
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import os
import shutil

def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text_content = ""
        tables_with_titles = []
        for page in pdf.pages:
            text_content += page.extract_text()
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    # Extracting the title before the table
                    table_start = table[0][1] if table else 0
                    text_before_table = page.extract_text()
                    lines = text_before_table.split('\n')
                    title = ""
                    for line in reversed(lines):
                        if line.strip():
                            title = line
                            break
                    tables_with_titles.append({"Title": title, "Table": table})
    return text_content, tables_with_titles


def select_pdf():
    global pdf_file
    pdf_file = filedialog.askopenfilename()
    if pdf_file:
        pdf_label.config(text=f"Selected PDF: {pdf_file}")
        pdf_label.pack()


def select_sample_excel():
    global sample_excel, sample_label
    sample_excel = filedialog.askopenfilename()
    if sample_excel:
        sample_label.config(text=f"Selected Sample Excel: {sample_excel}")
        sample_label.pack()


def generate_excel():
    global pdf_file, sample_excel
    if pdf_file:
        text_content, tables_with_titles = extract_data_from_pdf(pdf_file)
        
        output_dir = filedialog.askdirectory()
        print(output_dir)  # Debug line to check the selected directory
        
        if output_dir:
            temp_folder = os.path.join(output_dir, "temp118121")
            if not os.path.exists(temp_folder):
                os.makedirs(temp_folder)  # Create the 'temp118121' folder if it doesn't exist
            
            # Create files inside 'tempo118121' from PDF data
            if tables_with_titles:
                for idx, data in enumerate(tables_with_titles):
                    table = data["Table"]
                    df = pd.DataFrame(table)
                    df.ffill(axis=0, inplace=True)  # Fill empty cells with previous values in the same column
                    
                    excel_path = f"{temp_folder}/table_{idx}.xlsx"  # Using table number for filename within the temp folder
                    try:
                        df.to_excel(excel_path, index=False)
                        print(f"table_{idx}.xlsx created")  # Print the Excel file created
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to create Excel file at {excel_path}: {e}")
                        print(f"Failed to create Excel file at {excel_path}: {e}")  # Detailed error message in console
                
                messagebox.showinfo("Excel Created Successfully", f"Excel(s) created successfully in {temp_folder}!")
                
                # Proceed with the first Excel file
                process_first_excel(output_dir, tables_with_titles)
                clear_labels()  # Reset the labels after generating Excel files
                refresh_app()  # Refresh the app after successful task
            else:
                messagebox.showwarning("No Tables Found", "No tables were detected in the PDF.")
        else:
            messagebox.showwarning("Debug", "No output directory selected.")


def process_first_excel(output_dir, tables_with_titles):
    temp_folder = os.path.join(output_dir, "temp118121")
    if tables_with_titles:
        first_table_df = pd.DataFrame(tables_with_titles[0]["Table"])
        first_table_df.ffill(axis=0, inplace=True)

        if sample_excel:
            for row_idx, row in first_table_df.iloc[1:-1].iterrows():  # Exclude the first and last rows
                new_file_name = f"{row[0]}_{row[1]}.xlsx"  # Naming convention based on first and second column values
                shutil.copy(sample_excel, os.path.join(output_dir, new_file_name))

        messagebox.showinfo("Processing Completed", f"New Excel files created based on the first table!")
    else:
        messagebox.showwarning("No Tables Found", "No tables were detected in the PDF.")



def clear_labels():
    pdf_label.config(text="Selected PDF: ")
    pdf_label.pack()


def refresh_app():
    root.destroy()
    main()


def main():
    global root, pdf_label, sample_label

    root = tk.Tk()
    root.title("PDF to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_button = ttk.Button(main_frame, text="Select PDF", command=select_pdf, style="TButton")
    select_button.pack(pady=10)

    pdf_label = tk.Label(main_frame, text="Selected PDF: ", bg="#f0f0f0")
    pdf_label.pack()

    select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=select_sample_excel, style="TButton")
    select_sample_button.pack(pady=10)

    sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
    sample_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Excel", command=generate_excel, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
```
## Match First & Second table attributes
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import os
import shutil
import xlsxwriter
import openpyxl
import xlrd
import time



def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text_content = ""
        tables_with_titles = []
        for page in pdf.pages:
            text_content += page.extract_text()
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    # Extracting the title before the table
                    table_start = table[0][1] if table else 0
                    text_before_table = page.extract_text()
                    lines = text_before_table.split('\n')
                    title = ""
                    for line in reversed(lines):
                        if line.strip():
                            title = line
                            break
                    tables_with_titles.append({"Title": title, "Table": table})
    return text_content, tables_with_titles


def select_pdf():
    global pdf_file
    pdf_file = filedialog.askopenfilename()
    if pdf_file:
        pdf_label.config(text=f"Selected PDF: {pdf_file}")
        pdf_label.pack()


def select_sample_excel():
    global sample_excel, sample_label
    sample_excel = filedialog.askopenfilename()
    if sample_excel:
        sample_label.config(text=f"Selected Sample Excel: {sample_excel}")
        sample_label.pack()


def generate_excel():
    global pdf_file, sample_excel,output_dir
    if pdf_file:
        text_content, tables_with_titles = extract_data_from_pdf(pdf_file)
        
        output_dir = filedialog.askdirectory()
        print("Output directory is: "+output_dir)  # Debug line to check the selected directory
        
        if output_dir:
            temp_folder = os.path.join(output_dir, "temp118121")
            if not os.path.exists(temp_folder):
                os.makedirs(temp_folder)  # Create the 'temp118121' folder if it doesn't exist
            
            # Create files inside 'tempo118121' from PDF data
            if tables_with_titles:
                for idx, data in enumerate(tables_with_titles):
                    table = data["Table"]
                    df = pd.DataFrame(table)
                    df.ffill(axis=0, inplace=True)  # Fill empty cells with previous values in the same column
                    
                    excel_path = f"{temp_folder}/table_{idx}.xlsx"  # Using table number for filename within the temp folder
                    try:
                        df.to_excel(excel_path, index=False)
                        print(f"table_{idx}.xlsx created")  # Print the Excel file created
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to create Excel file at {excel_path}: {e}")
                        print(f"Failed to create Excel file at {excel_path}: {e}")  # Detailed error message in console
                
                messagebox.showinfo("Excel Created Successfully", f"Excel(s) created successfully in {temp_folder}!")
                
                # Proceed with the first Excel file
                process_first_excel(output_dir, temp_folder, tables_with_titles)
                clear_labels()  # Reset the labels after generating Excel files
                refresh_app()  # Refresh the app after successful task
            else:
                messagebox.showwarning("No Tables Found", "No tables were detected in the PDF.")
        else:
            messagebox.showwarning("Debug", "No output directory selected.")





def process_first_excel(output_dir, temp_folder, tables_with_titles):
    # temp_folder = os.path.join(output_dir, "temp118121")
    if len(tables_with_titles) >= 2:
        # Load data from the first two Excel files
        first_excel_path = os.path.join(temp_folder, "table_0.xlsx")
        second_excel_path = os.path.join(temp_folder, "table_1.xlsx")

        # Load first and second tables
        first_table_df = pd.read_excel(first_excel_path)
        second_table_df = pd.read_excel(second_excel_path)

        # Extract second column (except first and last rows) from the first table
        second_column_first_table = first_table_df.iloc[1:-1, 1]  # Assuming second column index is 1

        # Copy sample Excel to a new file named after values in the second column of the first table
        for temp_value in second_column_first_table:
            value=str(temp_value).replace(" ", "").replace(".", "").replace(",", "")
            new_file_name = f"{value}.xlsx"
            new_excel_file_path = os.path.join(output_dir, new_file_name)
            shutil.copy(sample_excel, os.path.join(output_dir, new_file_name))
            print(f"{new_file_name} created")  # Print the Excel file created

            # Check if any cell value in the second table matches the values in the second column of the first table
            for index, row in second_table_df.iterrows():
                cleaned_row_value = str(row[1]).replace(" ", "").replace(".", "").replace(",", "")  # Assuming the second column index is 1
                cleaned_value = str(value).replace(" ", "").replace(".", "").replace(",", "")  # Clean the value from the first table
                # print("Hand: "+ row[1], {value})
                print("Hand: "+ cleaned_row_value, cleaned_value)
                print(cleaned_row_value == cleaned_value)
                if cleaned_row_value == cleaned_value:  # Assuming the comparison column in the second table is index 1
                    retrieved_value = row[3]  # Assuming the retrieved column in the second table is index 3
                    print("Match Found")
                    # Write the retrieved value to the new Excel file in row 27, column 7
                    new_excel_file_path = os.path.join(output_dir, new_file_name)
                    print(f"New Excel file path: {new_excel_file_path}")

                   # Code here
    else:
        messagebox.showwarning("No Tables Found", "Insufficient tables detected in the PDF.")




def clear_labels():
    pdf_label.config(text="Selected PDF: ")
    pdf_label.pack()


def refresh_app():
    root.destroy()
    main()


def main():
    global root, pdf_label, sample_label

    root = tk.Tk()
    root.title("PDF to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_button = ttk.Button(main_frame, text="Select PDF", command=select_pdf, style="TButton")
    select_button.pack(pady=10)

    pdf_label = tk.Label(main_frame, text="Selected PDF: ", bg="#f0f0f0")
    pdf_label.pack()

    select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=select_sample_excel, style="TButton")
    select_sample_button.pack(pady=10)

    sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
    sample_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Excel", command=generate_excel, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
```
## last task till 28/11/2023
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import os
import shutil
from openpyxl import load_workbook




def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text_content = ""
        tables_with_titles = []
        for page in pdf.pages:
            text_content += page.extract_text()
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    # Extracting the title before the table
                    table_start = table[0][1] if table else 0
                    text_before_table = page.extract_text()
                    lines = text_before_table.split('\n')
                    title = ""
                    for line in reversed(lines):
                        if line.strip():
                            title = line
                            break
                    tables_with_titles.append({"Title": title, "Table": table})
    return text_content, tables_with_titles


def select_pdf():
    global pdf_file
    pdf_file = filedialog.askopenfilename()
    if pdf_file:
        pdf_label.config(text=f"Selected PDF: {pdf_file}")
        pdf_label.pack()


def select_sample_excel():
    global sample_excel, sample_label
    sample_excel = filedialog.askopenfilename()
    if sample_excel:
        sample_label.config(text=f"Selected Sample Excel: {sample_excel}")
        sample_label.pack()


def generate_excel():
    global pdf_file, sample_excel,output_dir
    if pdf_file:
        text_content, tables_with_titles = extract_data_from_pdf(pdf_file)
        
        output_dir = filedialog.askdirectory()
        print("Output directory is: "+output_dir)  # Debug line to check the selected directory
        
        if output_dir:
            temp_folder = os.path.join(output_dir, "temp118121")
            if not os.path.exists(temp_folder):
                os.makedirs(temp_folder)  # Create the 'temp118121' folder if it doesn't exist
            
            # Create files inside 'tempo118121' from PDF data
            if tables_with_titles:
                for idx, data in enumerate(tables_with_titles):
                    table = data["Table"]
                    df = pd.DataFrame(table)
                    df.ffill(axis=0, inplace=True)  # Fill empty cells with previous values in the same column
                    
                    excel_path = f"{temp_folder}/table_{idx}.xlsx"  # Using table number for filename within the temp folder
                    try:
                        df.to_excel(excel_path, index=False)
                        print(f"table_{idx}.xlsx created")  # Print the Excel file created
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to create Excel file at {excel_path}: {e}")
                        print(f"Failed to create Excel file at {excel_path}: {e}")  # Detailed error message in console
                
                messagebox.showinfo("Excel Created Successfully", f"Excel(s) created successfully in {temp_folder}!")
                
                # Proceed with the first Excel file
                process_first_excel(output_dir, tables_with_titles)
                clear_labels()  # Reset the labels after generating Excel files
                refresh_app()  # Refresh the app after successful task
            else:
                messagebox.showwarning("No Tables Found", "No tables were detected in the PDF.")
        else:
            messagebox.showwarning("Debug", "No output directory selected.")





def process_first_excel(output_dir, tables_with_titles):
    temp_folder = os.path.join(output_dir, "temp118121")
    sample_file_name = os.path.basename(sample_excel)  # Extracting the sample file name

    if tables_with_titles and sample_excel:
        for idx in range(2):  # Considering only the first two tables for processing
            table_path = os.path.join(temp_folder, f"table_{idx}.xlsx")
            if os.path.exists(table_path):
                table_df = pd.read_excel(table_path)
                second_column_first_table = table_df.iloc[1:-1, 1]

                for temp_value in second_column_first_table:
                    value = str(temp_value).replace(' ', '').replace('.', '').replace(',', '')  # Clean the value
                    new_file_name = f"{value}.xlsx"
                    new_file_path = os.path.join(output_dir, new_file_name)

                    shutil.copy(sample_excel, new_file_path)
                    print("File Created", f"New file '{new_file_name}' created.")

                    for next_idx in range(idx + 1, len(tables_with_titles)):
                        next_table_path = os.path.join(temp_folder, f"table_{next_idx}.xlsx")
                        if os.path.exists(next_table_path):
                            next_table_df = pd.read_excel(next_table_path)
                            second_column_next_table = next_table_df.iloc[1:, 1]

                            for match_name in second_column_next_table:
                                match_name = str(match_name).replace(' ', '').replace('.', '').replace(',', '')
                                matches = [file for file in os.listdir(output_dir) if match_name in file]
                                if matches:
                                    print("Match Found", "Matched")

                                    for match in matches:
                                        file_path = os.path.join(output_dir, match)
                                        matched_excel = pd.read_excel(file_path)
                                        for idx, m_row in matched_excel.iterrows():
                                            if value in str(m_row.iloc[0]):
                                                retrieved_value = m_row.iloc[3]
                                                print("Match Found", f"Match found with value: {retrieved_value}")

                                                workbook = load_workbook(file_path)
                                                sheet = workbook.active
                                                sheet['G27'] = retrieved_value
                                                workbook.save(file_path)
                                                messagebox.showinfo("Value Written", f"Value '{retrieved_value}' written to '{match}'.")

        messagebox.showinfo("Process Completed", "All files updated with matches.")





def clear_labels():
    pdf_label.config(text="Selected PDF: ")
    pdf_label.pack()


def refresh_app():
    root.destroy()
    main()


def main():
    global root, pdf_label, sample_label

    root = tk.Tk()
    root.title("PDF to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_button = ttk.Button(main_frame, text="Select PDF", command=select_pdf, style="TButton")
    select_button.pack(pady=10)

    pdf_label = tk.Label(main_frame, text="Selected PDF: ", bg="#f0f0f0")
    pdf_label.pack()

    select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=select_sample_excel, style="TButton")
    select_sample_button.pack(pady=10)

    sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
    sample_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Excel", command=generate_excel, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
```
# Working With Doc FIle
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
from docx import Document

# Function to extract data from a Word document
def extract_data_from_docx(docx_file):
    doc = Document(docx_file)
    text_content = ""
    tables_with_titles = []
    
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text)
            table_data.append(row_data)
        
        # Extracting the title before the table
        title = ""
        for paragraph in table.rows[0].cells[0].paragraphs:
            title += paragraph.text
        tables_with_titles.append({"Title": title, "Table": table_data})
    
    return text_content, tables_with_titles

# Function to handle selection of Word document
def select_docx():
    global docx_file
    docx_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if docx_file:
        docx_label.config(text=f"Selected Word Doc: {docx_file}")
        docx_label.pack()

# Function to handle selection of Sample Excel file
def select_sample_excel():
    global sample_excel, sample_label
    sample_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if sample_excel:
        sample_label.config(text=f"Selected Sample Excel: {sample_excel}")
        sample_label.pack()


# Function to handle generation of Excel from Word document
def generate_excel_from_docx():
    global docx_file, sample_excel
    if docx_file:
        text_content, tables_with_titles = extract_data_from_docx(docx_file)
        
        output_dir = filedialog.askdirectory()
        
        if output_dir:
            temp_folder = os.path.join(output_dir, "temp118121")
            if not os.path.exists(temp_folder):
                os.makedirs(temp_folder)
            
            if tables_with_titles:
                for idx, data in enumerate(tables_with_titles):
                    table = data["Table"]
                    df = pd.DataFrame(table)
                    df.ffill(axis=0, inplace=True)
                    
                    excel_path = f"{output_dir}/table_{idx}.xlsx"
                    try:
                        df.to_excel(excel_path, index=False)
                        print(f"table_{idx}.xlsx created")
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to create Excel file at {excel_path}: {e}")
                        print(f"Failed to create Excel file at {excel_path}: {e}")
                    
                    # Move generated Excel files to temp folder
                    if os.path.exists(excel_path):
                        shutil.move(excel_path, os.path.join(temp_folder, f"table_{idx}.xlsx"))
                
                messagebox.showinfo("Excel Created Successfully", f"Excel(s) moved to {temp_folder}!")
                process_first_excel(output_dir, tables_with_titles)
                clear_labels()
            else:
                messagebox.showwarning("No Tables Found", "No tables were detected in the Word document.")
        else:
            messagebox.showwarning("Debug", "No output directory selected.")


def process_first_excel(output_dir, tables_with_titles):
    if tables_with_titles and sample_excel:
        first_table_df = pd.DataFrame(tables_with_titles[0]["Table"])
        first_table_df.ffill(axis=0, inplace=True)

        for row_idx, row in first_table_df.iloc[1:-1].iterrows():
            new_file_name = f"{row[1]}.xlsx"
            shutil.copy(sample_excel, os.path.join(output_dir, new_file_name))

        messagebox.showinfo("Processing Completed", f"New Excel files created based on the first table!")
    else:
        messagebox.showwarning("No Tables Found or No Sample Excel", "No tables were detected in the Word document or no Sample Excel selected.")

def clear_labels():
    docx_label.config(text="Selected Word Doc: ")
    docx_label.pack()
    sample_label.config(text="Selected Sample Excel: ")
    sample_label.pack()

def main():
    global root, docx_label, sample_label, sample_excel

    root = tk.Tk()
    root.title("Word to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_button = ttk.Button(main_frame, text="Select Word Doc", command=select_docx, style="TButton")
    select_button.pack(pady=10)

    docx_label = tk.Label(main_frame, text="Selected Word Doc: ", bg="#f0f0f0")
    docx_label.pack()

    select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=select_sample_excel, style="TButton")
    select_sample_button.pack(pady=10)

    sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
    sample_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Table in Excel", command=generate_excel_from_docx, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
```
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
from docx import Document


# Function to handle selection of Word document
def select_docx():
    global docx_file
    docx_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if docx_file:
        docx_label.config(text=f"Selected Word Doc: {docx_file}")
        docx_label.pack()

# Function to handle selection of Sample Excel file
def select_sample_excel():
    global sample_excel, sample_label
    sample_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if sample_excel:
        sample_label.config(text=f"Selected Sample Excel: {sample_excel}")
        sample_label.pack()


# Function to extract data from a Word document
def extract_data_from_docx(docx_file):
    doc = Document(docx_file)
    text_content = ""
    tables_with_titles = []
    
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text)
            table_data.append(row_data)
        
        # Extracting the title before the table
        title = ""
        for paragraph in table.rows[0].cells[0].paragraphs:
            title += paragraph.text
        tables_with_titles.append({"Title": title, "Table": table_data})
    
    return text_content, tables_with_titles


# Function to handle generation of Excel from Word document
def generate_excel_from_docx():
    global docx_file, sample_excel
    if docx_file:
        text_content, tables_with_titles = extract_data_from_docx(docx_file)
        
        output_dir = filedialog.askdirectory()
        
        if output_dir:
            temp_folder = os.path.join(output_dir, "temp118121")
            if not os.path.exists(temp_folder):
                os.makedirs(temp_folder)
            
            if tables_with_titles:
                for idx, data in enumerate(tables_with_titles):
                    table = data["Table"]
                    df = pd.DataFrame(table)
                    df.ffill(axis=0, inplace=True)
                    
                    excel_path = f"{temp_folder}/table_{idx}.xlsx"
                    try:
                        df.to_excel(excel_path, index=False)
                        print(f"table_{idx}.xlsx created")
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to create Excel file at {excel_path}: {e}")
                        print(f"Failed to create Excel file at {excel_path}: {e}")
                
                messagebox.showinfo("Excel Created Successfully", f"Excel(s) moved to {temp_folder}!")
                process_first_excel(output_dir, tables_with_titles)
                clear_labels()
            else:
                messagebox.showwarning("No Tables Found", "No tables were detected in the Word document.")
        else:
            messagebox.showwarning("Debug", "No output directory selected.")
    else:
        messagebox.showwarning("Opps!", "Please select valid doc file.")


def process_first_excel(output_dir, tables_with_titles):
    if tables_with_titles and sample_excel:
        first_table_df = pd.DataFrame(tables_with_titles[0]["Table"])
        first_table_df.ffill(axis=0, inplace=True)
        created_files = []

        for row_idx, row in first_table_df.iloc[1:-1].iterrows():
            new_file_name = row[1].replace(" ", "").replace(".", "").replace(",", "") + ".xlsx"
            print(f"Creating {new_file_name}...")
            shutil.copy(sample_excel, os.path.join(output_dir, new_file_name))
            created_files.append(new_file_name)

        messagebox.showinfo("Processing Completed", f"New Excel files created based on the first table!")

        # Check for existing files
        # existing_files = [filename for filename in os.listdir(output_dir) if filename.endswith(".xlsx")]
        # duplicates = set(created_files) & set(existing_files)
        # if duplicates:
        #     messagebox.showwarning("Duplicates Found", f"Duplicate files found: {', '.join(duplicates)}")
    else:
        messagebox.showwarning("No Tables Found or No Sample Excel", "No tables were detected in the Word document or no Sample Excel selected.")
def clear_labels():
    docx_label.config(text="Selected Word Doc: ")
    docx_label.pack()
    sample_label.config(text="Selected Sample Excel: ")
    sample_label.pack()

def main():
    global root, docx_label, sample_label, sample_excel

    root = tk.Tk()
    root.title("Word to Excel Converter")

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
    style.map("TButton", background=[("active", "#0056b3")])

    main_frame = tk.Frame(root, bg="#f0f0f0")
    main_frame.pack(padx=20, pady=20)

    select_button = ttk.Button(main_frame, text="Select Word Doc", command=select_docx, style="TButton")
    select_button.pack(pady=10)

    docx_label = tk.Label(main_frame, text="Selected Word Doc: ", bg="#f0f0f0")
    docx_label.pack()

    select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=select_sample_excel, style="TButton")
    select_sample_button.pack(pady=10)

    sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
    sample_label.pack()

    generate_button = ttk.Button(main_frame, text="Generate Table in Excel", command=generate_excel_from_docx, style="TButton")
    generate_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
```

# Second table data entry successfull

```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
from docx import Document
from openpyxl import load_workbook
import xlwings as xw


class WordToExcelConverter:
    def __init__(self):
        self.docx_file = None
        self.sample_excel = None
        self.output_dir = None
        self.tables_with_titles = None
        self.combined_excel_path = None
        self.new_files = []  # Array to store new_file values globally
    
    def clear_labels(self):
        self.docx_label.config(text="Selected Word Doc: ")
        self.docx_label.pack()
        self.sample_label.config(text="Selected Sample Excel: ")
        self.sample_label.pack()

    def select_docx(self):
        # Function to handle selection of Word document
        self.docx_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if self.docx_file:
            self.docx_label.config(text=f"Selected Word Doc: {self.docx_file}")
            self.docx_label.pack()

    def select_sample_excel(self):
        # Function to handle selection of Sample Excel file
        self.sample_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.sample_excel:
            self.sample_label.config(text=f"Selected Sample Excel: {self.sample_excel}")
            self.sample_label.pack()

    def extract_data_from_docx(self):
        # Function to extract data from a Word document
        doc = Document(self.docx_file)
        text_content = ""
        self.tables_with_titles = []
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                table_data.append(row_data)
            # Extracting the title before the table
            title = ""
            for paragraph in table.rows[0].cells[0].paragraphs:
                title += paragraph.text
            self.tables_with_titles.append({"Title": title, "Table": table_data})
        return text_content, self.tables_with_titles

    def generate_excel_from_docx(self):
        if self.docx_file:
            text_content, self.tables_with_titles = self.extract_data_from_docx()
            self.output_dir = filedialog.askdirectory()
            temp_folder = os.path.join(self.output_dir, "temp118121")  # Path to temp118121 directory
            if self.output_dir:
                if not os.path.exists(temp_folder):
                    os.makedirs(temp_folder)  # Create temp118121 directory if it doesn't exist
                
                if self.tables_with_titles:
                    excel_path = os.path.join(temp_folder, "tables_combined.xlsx")  # Save tables_combined.xlsx inside temp118121
                    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                        for i, data in enumerate(self.tables_with_titles):
                            table = data["Table"]
                            df = pd.DataFrame(table)
                            df.ffill(axis=0, inplace=True)
                            sheet_name = f"Table_{i}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"{sheet_name} added to Excel")
                        messagebox.showinfo("Excel Created Successfully", f"All tables moved to {excel_path}!")
                        # self.clear_labels()
                else:
                    messagebox.showwarning("No Tables Found", "No tables were detected in the Word document.")
            else:
                messagebox.showwarning("Debug", "No output directory selected.")
        else:
            messagebox.showwarning("Oops!", "Please select a valid doc file.")


    def process_first_excel(self):
        if self.output_dir and self.tables_with_titles and self.sample_excel:
            file_count = 0  # Counter for the files being created
            combined_df = pd.DataFrame()  # Initialize an empty DataFrame to hold all tables

            for i, data in enumerate(self.tables_with_titles):
                table = data["Table"]
                df = pd.DataFrame(table)
                df.ffill(axis=0, inplace=True)
                if i == 0:  # Working with the first table
                    first_table_df = df.iloc[:, 1]  # Extracting the content from the second column

                    # Create separate Excel files based on each row's content
                    for row_i, value in enumerate(first_table_df):
                        if row_i != 0 and row_i != len(first_table_df) - 1:
                            new_file=value.replace(" ", "").replace(".", "").replace(",", "")
                            self.new_files.append(new_file)  # Append new_file to the global array
                            new_file_name = new_file + ".xlsx"
                            print(f"Creating {new_file_name}...")
                            shutil.copy(self.sample_excel, os.path.join(self.output_dir, new_file_name))
                            file_count += 1  # Increment file count

                    # Append the first table content to the combined DataFrame
                    combined_df = pd.concat([combined_df, first_table_df], axis=1)
                else:
                    combined_df = pd.concat([combined_df, df.iloc[1:-1, 1]], axis=1)
            print(file_count)
            print(self.new_files)
            messagebox.showinfo("Congratulations!", f"Excel Created Successfully! Total Files Created: {file_count}")
            # self.clear_labels()
        else:
            messagebox.showwarning("No Tables Found or No Sample Excel", "No tables were detected in the Word document or no Sample Excel selected.")


    def print_matching_value(self):
        # An array of the size of the first, initially all value is 0
        matching_values = [0] * len(self.new_files)

        # Check for substring matches in the second table of tables_combined
        if len(self.tables_with_titles) >= 2:
            second_table_data = self.tables_with_titles[1]["Table"]
            second_table_df = pd.DataFrame(second_table_data)

            for i, file_value in enumerate(self.new_files):
                # Check if the file_value is a substring of any value in the 2nd column of the second table
                for row_i in range(1, len(second_table_df)):  # Start from the second row
                    table_value = str(second_table_df.iloc[row_i, 1]).replace(" ", "").replace(".", "").replace(",", "")
                    print(file_value, " ", table_value)
                    if str(file_value).lower() in table_value.lower() or table_value.lower() in str(file_value).lower():
                        matching_values[i] += float(second_table_df.iloc[row_i, 3])  # Take the value from the 4th column
                        print("match")
        for i in range(1,len(matching_values)):
            print(f"{self.new_files[i] , matching_values[i]}")

        for i, file_name in enumerate(self.new_files):
            if i > 0 and matching_values[i] != 0:
                file_path = os.path.join(self.output_dir, f"{file_name}.xlsx")
                if os.path.exists(file_path):
                    try:
                        print("Inserting data at ", f"{file_name}.xlsx")
                        wb = xw.Book(file_path)
                        sheet = wb.sheets.active
                        sheet.range('G9').value = matching_values[i]
                        wb.save(file_path)
                        wb.close()
                    except Exception as e:
                        print(f"Error processing file {file_path}: {e}")
                else:
                    print(f"File {file_path} does not exist.")


        messagebox.showinfo("Your task is successfull!", f"Output Excel File Created Successfully!")
        self.clear_labels()



    def main(self):
        root = tk.Tk()
        root.title("Word to Excel Converter")
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
        style.map("TButton", background=[("active", "#0056b3")])
        main_frame = tk.Frame(root, bg="#f0f0f0")
        main_frame.pack(padx=20, pady=20)

        select_button = ttk.Button(main_frame, text="Select Word Doc", command=self.select_docx, style="TButton")
        select_button.pack(pady=10)
        self.docx_label = tk.Label(main_frame, text="Selected Word Doc: ", bg="#f0f0f0")
        self.docx_label.pack()

        select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=self.select_sample_excel, style="TButton")
        select_sample_button.pack(pady=10)
        self.sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
        self.sample_label.pack()

        generate_button = ttk.Button(main_frame, text="Generate Table in Excel", command=self.generate_excel_from_docx, style="TButton")
        generate_button.pack(pady=10)

        process_button = ttk.Button(main_frame, text="Process the first table", command=self.process_first_excel, style="TButton")
        process_button.pack(pady=10)

        check_matching_button = ttk.Button(main_frame, text="Check Matching Values", command=self.print_matching_value, style="TButton")
        check_matching_button.pack(pady=10)


        root.mainloop()


if __name__ == "__main__":
    converter = WordToExcelConverter()
    converter.main()
    # messagebox.showinfo("Your task is successfull!", f"Output Excel File Created Successfully!")
```
# Almost Done
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
from docx import Document
from openpyxl import load_workbook
import xlwings as xw


class WordToExcelConverter:
    def __init__(self):
        self.docx_file = None
        self.sample_excel = None
        self.output_dir = None
        self.tables_with_titles = None
        self.combined_excel_path = None
        self.new_files = []  # Array to store new_file values globally
        self.total_no_of_table=0
    
    def clear_labels(self):
        self.docx_label.config(text="Selected Word Doc: ")
        self.docx_label.pack()
        self.sample_label.config(text="Selected Sample Excel: ")
        self.sample_label.pack()

    def select_docx(self):
        # Function to handle selection of Word document
        self.docx_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if self.docx_file:
            self.docx_label.config(text=f"Selected Word Doc: {self.docx_file}")
            self.docx_label.pack()

    def select_sample_excel(self):
        # Function to handle selection of Sample Excel file
        self.sample_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.sample_excel:
            self.sample_label.config(text=f"Selected Sample Excel: {self.sample_excel}")
            self.sample_label.pack()

    def extract_data_from_docx(self):
        # Function to extract data from a Word document
        doc = Document(self.docx_file)
        text_content = ""
        self.tables_with_titles = []
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                table_data.append(row_data)
            # Extracting the title before the table
            title = ""
            for paragraph in table.rows[0].cells[0].paragraphs:
                title += paragraph.text
            self.tables_with_titles.append({"Title": title, "Table": table_data})
            self.total_no_of_table = len(self.tables_with_titles)
        return text_content, self.tables_with_titles

    def generate_excel_from_docx(self):
        if self.docx_file:
            text_content, self.tables_with_titles = self.extract_data_from_docx()
            self.output_dir = filedialog.askdirectory()
            temp_folder = os.path.join(self.output_dir, "temp118121")  # Path to temp118121 directory
            if self.output_dir:
                if not os.path.exists(temp_folder):
                    os.makedirs(temp_folder)  # Create temp118121 directory if it doesn't exist
                
                if self.tables_with_titles:
                    excel_path = os.path.join(temp_folder, "tables_combined.xlsx")  # Save tables_combined.xlsx inside temp118121
                    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                        for i, data in enumerate(self.tables_with_titles):
                            table = data["Table"]
                            df = pd.DataFrame(table)
                            df.ffill(axis=0, inplace=True)
                            sheet_name = f"Table_{i}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"{sheet_name} added to Excel")
                        messagebox.showinfo("Excel Created Successfully", f"All tables moved to {excel_path}!")
                        # self.clear_labels()
                else:
                    messagebox.showwarning("No Tables Found", "No tables were detected in the Word document.")
            else:
                messagebox.showwarning("Debug", "No output directory selected.")
        else:
            messagebox.showwarning("Oops!", "Please select a valid doc file.")


    def print_matching_value_for_file(self, new_file):
        print("Processing...")
        total_no_of_table = len(self.tables_with_titles) #12

        # An array of the size of the first, initially all value is 0
        matching_values = [0] * total_no_of_table #12

        # Question Paper Setter & Script Examiner 
        if total_no_of_table > 1:
            second_table_data = self.tables_with_titles[1]["Table"]
            second_table_df = pd.DataFrame(second_table_data)

            for row_idx in range(1, len(second_table_df)):
                table_value = str(second_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[1] += float(second_table_df.iloc[row_idx, 3]) 
                    print(f"Matching value for {new_file}: {matching_values[1]}")

        # Examiners of Class Tests
        if total_no_of_table > 2:
            third_table_data = self.tables_with_titles[2]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[2] += float(third_table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[2]}")

        # Examiners of Sessional Classes
        if total_no_of_table > 3:
            third_table_data = self.tables_with_titles[3]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[3] += float(third_table_df.iloc[row_idx, 2])*float(third_table_df.iloc[row_idx, 3])/1.5
                    print(f"Matching value for {new_file}: {matching_values[3]}")

        # # Script Scrutinizer
        # if total_no_of_table > 4:
        #     third_table_data = self.tables_with_titles[4]["Table"]
        #     third_table_df = pd.DataFrame(third_table_data)

        #     for row_idx in range(1, len(third_table_df)):
        #         table_value = str(third_table_df.iloc[row_idx, 0]).replace(" ", "").replace(".", "").replace(",", "")
        #         if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
        #             matching_values[4] += float(third_table_df.iloc[row_idx, 1]) 
        #             print(f"Matching value for {new_file}: {matching_values[4]}")


        # Tabulation & Verification
        if total_no_of_table > 5:
            third_table_data = self.tables_with_titles[5]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[5] += float(third_table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[5]}")


        # Typing and Drawing
        if total_no_of_table > 6:
            third_table_data = self.tables_with_titles[6]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 0]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[6] += float(third_table_df.iloc[row_idx, 1]) 
                    print(f"Matching value for {new_file}: {matching_values[6]}")


        # Central Viva-Voce
        if total_no_of_table > 7:
            third_table_data = self.tables_with_titles[7]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[7] += float(third_table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[7]}")


        # Student Advising
        if total_no_of_table > 8:
            third_table_data = self.tables_with_titles[8]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[8] += float(third_table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[8]}")


        # Seminar (CSE 4120) 1 + 1 =2
        if total_no_of_table > 9:
            third_table_data = self.tables_with_titles[9]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[9] += float(third_table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[9]}")


        # Thesis Progress Defense
        if total_no_of_table > 10:
            third_table_data = self.tables_with_titles[10]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[10] += float(third_table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[10]}")


        # # Final Grade Sheet Verification
        # if total_no_of_table > 11:
        #     third_table_data = self.tables_with_titles[11]["Table"]
        #     third_table_df = pd.DataFrame(third_table_data)

        #     for row_idx in range(1, len(third_table_df)):
        #         table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
        #         if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
        #             matching_values[11] += float(third_table_df.iloc[row_idx, 2]) 
        #             print(f"Matching value for {new_file}: {matching_values[11]}")


        
        # Thesis Progress Defense
        if total_no_of_table > 12:
            third_table_data = self.tables_with_titles[12]["Table"]
            third_table_df = pd.DataFrame(third_table_data)

            for row_idx in range(1, len(third_table_df)):
                table_value = str(third_table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[12] += float(third_table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[12]}")




        file_path = os.path.join(self.output_dir, f"{new_file}.xlsx")
        cell_mappings = {
            1: ['G9', 'G12'],
            2: ['G14'],
            3: ['G17'],
            4: ['G25'],
            5: ['G23', 'G24'],
            6: ['G27'],
            7: ['G18'],
            8: ['G29'],
            9: ['G16'],
            10: ['G20'],
            11: ['G28'],
            12: ['G26']
        }

        for i in range(1, total_no_of_table):
            if os.path.exists(file_path):
                try:
                    if matching_values[i] != 0:
                        print("Inserting data at ", f"{new_file}.xlsx")
                        wb = xw.Book(file_path)
                        sheet = wb.sheets.active
                        cell_locations = cell_mappings.get(i, [])
                        for cell in cell_locations:
                            sheet.range(cell).value = matching_values[i]
                        wb.save(file_path)
                        wb.close()
                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")
            else:
                print(f"File {file_path} does not exist.")




        print(matching_values)
        



    def process_first_excel(self):
        if self.output_dir and self.tables_with_titles and self.sample_excel:
            file_count = 0  # Counter for the files being created
            combined_df = pd.DataFrame()  # Initialize an empty DataFrame to hold all tables

            for i, data in enumerate(self.tables_with_titles):
                table = data["Table"]
                df = pd.DataFrame(table)
                df.ffill(axis=0, inplace=True)
                if i == 0:  # Working with the first table
                    first_table_df = df.iloc[:, 1]  # Extracting the content from the second column

                    # Create separate Excel files based on each row's content
                    for row_i, value in enumerate(first_table_df):
                        if row_i != 0 and row_i != len(first_table_df) - 1:
                            new_file=value.replace(" ", "").replace(".", "").replace(",", "")
                            self.new_files.append(new_file)  # Append new_file to the global array
                            new_file_name = new_file + ".xlsx"
                            print(f"Creating {new_file_name}...")
                            shutil.copy(self.sample_excel, os.path.join(self.output_dir, new_file_name))
                            self.print_matching_value_for_file(new_file)
                            file_count += 1  # Increment file count

                    # Append the first table content to the combined DataFrame
                    combined_df = pd.concat([combined_df, first_table_df], axis=1)
                else:
                    combined_df = pd.concat([combined_df, df.iloc[1:-1, 1]], axis=1)
            print("The total no of files are :",file_count)
            print("The files are: ",self.new_files)
            messagebox.showinfo("Congratulations!", f"Excel Created Successfully! Total Files Created: {file_count}")
            self.clear_labels()
        else:
            messagebox.showwarning("No Tables Found or No Sample Excel", "No tables were detected in the Word document or no Sample Excel selected.")




    def main(self):
        root = tk.Tk()
        root.title("Word to Excel Converter")
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
        style.map("TButton", background=[("active", "#0056b3")])
        main_frame = tk.Frame(root, bg="#f0f0f0")
        main_frame.pack(padx=20, pady=20)

        select_button = ttk.Button(main_frame, text="Select Word Doc", command=self.select_docx, style="TButton")
        select_button.pack(pady=10)
        self.docx_label = tk.Label(main_frame, text="Selected Word Doc: ", bg="#f0f0f0")
        self.docx_label.pack()

        select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=self.select_sample_excel, style="TButton")
        select_sample_button.pack(pady=10)
        self.sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
        self.sample_label.pack()

        generate_button = ttk.Button(main_frame, text="Generate Table in Excel", command=self.generate_excel_from_docx, style="TButton")
        generate_button.pack(pady=10)

        process_button = ttk.Button(main_frame, text="Process the first table", command=self.process_first_excel, style="TButton")
        process_button.pack(pady=10)



        root.mainloop()


if __name__ == "__main__":
    converter = WordToExcelConverter()
    converter.main()
```
```
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
from docx import Document
from openpyxl import load_workbook
import xlwings as xw
import re
from docx import Document



class WordToExcelConverter:

    
    def __init__(self):
        self.docx_file = None
        self.sample_excel = None
        self.output_dir = None
        self.tables_with_titles = None
        self.combined_excel_path = None
        self.new_files = []  # Array to store new_file values globally
        self.total_no_of_table=0
        self.year=0
        self.term=0
        self.dept="CSE"
    
    def clear_labels(self):
        self.docx_label.config(text="Selected Word Doc: ")
        self.docx_label.pack()
        self.sample_label.config(text="Selected Sample Excel: ")
        self.sample_label.pack()

    def select_docx(self):
        # Function to handle selection of Word document
        self.docx_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if self.docx_file:
            self.docx_label.config(text=f"Selected Word Doc: {self.docx_file}")
            self.docx_label.pack()

    def select_sample_excel(self):
        # Function to handle selection of Sample Excel file
        self.sample_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.sample_excel:
            self.sample_label.config(text=f"Selected Sample Excel: {self.sample_excel}")
            self.sample_label.pack()

    def extract_words_before_table(self):
        pattern = r'Bills - (\w+).*?year (\w+)'
        doc = Document(self.docx_file)
        for paragraph in doc.paragraphs:
            for table in doc.tables:
                if paragraph in table._element.iterancestors('w:tbl'):
                    return None, None  # Stop searching when a paragraph is part of a table

            match = re.search(pattern, paragraph.text)
            if match:
                self.year = match.group(1)
                self.term = match.group(2)
                # print("Year, Term: ",self.year, self.term)
                print("Year & Term extracted Successfully")
                # messagebox.showinfo("Year & Term extracted Successfully", f"Year & Term extracted Successfully!")
                return

        print("Year & Term extraction failed")
        # messagebox.showwarning("No Year & Term Found", "No Year & Term Found!")

    def extract_data_from_docx(self):
        # Function to extract data from a Word document

        try:
            self.extract_words_before_table()
    
            if self.year and self.term:
                print("Word after 'bills :", self.year)
                print("Word after 'year':", self.term)
            else:
                print("No match found before the table.")
        except FileNotFoundError:
            print(f"Error: The file '{self.docx_file}' was not found.")
        except Exception as e:
            print(f"Error: {e}")
        
        doc = Document(self.docx_file)
        text_content = ""
        self.tables_with_titles = []
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                table_data.append(row_data)
            # Extracting the title before the table
            title = ""
            for paragraph in table.rows[0].cells[0].paragraphs:
                title += paragraph.text
            self.tables_with_titles.append({"Title": title, "Table": table_data})
            self.total_no_of_table = len(self.tables_with_titles)
        return text_content, self.tables_with_titles

    def generate_excel_from_docx(self):
        if self.docx_file:
            text_content, self.tables_with_titles = self.extract_data_from_docx()
            self.output_dir = filedialog.askdirectory()
            temp_folder = os.path.join(self.output_dir, "temp118121")  # Path to temp118121 directory
            if self.output_dir:
                if not os.path.exists(temp_folder):
                    os.makedirs(temp_folder)  # Create temp118121 directory if it doesn't exist
                
                if self.tables_with_titles:
                    excel_path = os.path.join(temp_folder, "tables_combined.xlsx")  # Save tables_combined.xlsx inside temp118121
                    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                        for i, data in enumerate(self.tables_with_titles):
                            table = data["Table"]
                            df = pd.DataFrame(table)
                            df.ffill(axis=0, inplace=True)
                            sheet_name = f"Table_{i}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"{sheet_name} added to Excel")
                        messagebox.showinfo("Excel Created Successfully", f"All tables moved to {excel_path}!")
                        # self.clear_labels()
                else:
                    messagebox.showwarning("No Tables Found", "No tables were detected in the Word document.")
            else:
                messagebox.showwarning("Debug", "No output directory selected.")
        else:
            messagebox.showwarning("Oops!", "Please select a valid doc file.")


    def print_matching_value_for_file(self, new_file, name, designation):
        print("Processing...")
        total_no_of_table = len(self.tables_with_titles) #12

        # An array of the size of the first, initially all value is 0
        matching_values = [0] * total_no_of_table #12


        # Set Name, Year, Term
        print("Name: ",name)
        print("Designation: ",designation)
        print("Year, Term: ",self.year,self.term)


        # Question Paper Setter & Script Examiner 
        if total_no_of_table > 1:
            table_data = self.tables_with_titles[1]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[1] += float(table_df.iloc[row_idx, 3]) 
                    print(f"Matching value for {new_file}: {matching_values[1]}")

        # Examiners of Class Tests
        if total_no_of_table > 2:
            table_data = self.tables_with_titles[2]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[2] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[2]}")

        # Examiners of Sessional Classes
        if total_no_of_table > 3:
            table_data = self.tables_with_titles[3]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[3] += float(table_df.iloc[row_idx, 2])*float(table_df.iloc[row_idx, 3])/1.5
                    print(f"Matching value for {new_file}: {matching_values[3]}")

        # Script Scrutinizer
        if total_no_of_table > 4:
            table_data = self.tables_with_titles[4]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 0]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[4] += float(table_df.iloc[row_idx, 1]) 
                    print(f"Matching value for {new_file}: {matching_values[4]}")


        # Tabulation & Verification
        if total_no_of_table > 5:
            table_data = self.tables_with_titles[5]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[5] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[5]}")


        # Typing and Drawing
        if total_no_of_table > 6:
            table_data = self.tables_with_titles[6]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 0]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[6] += float(table_df.iloc[row_idx, 1]) 
                    print(f"Matching value for {new_file}: {matching_values[6]}")


        # Central Viva-Voce
        if total_no_of_table > 7:
            table_data = self.tables_with_titles[7]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[7] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[7]}")


        # Student Advising
        if total_no_of_table > 8:
            table_data = self.tables_with_titles[8]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[8] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[8]}")


        # Seminar (CSE 4120) 1 + 1 =2
        if total_no_of_table > 9:
            table_data = self.tables_with_titles[9]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[9] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[9]}")


        # Thesis Progress Defense
        if total_no_of_table > 10:
            table_data = self.tables_with_titles[10]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[10] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[10]}")


        # Final Grade Sheet Verification
        if total_no_of_table > 11:
            table_data = self.tables_with_titles[11]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    print(table_df)
                    matching_values[11] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[11]}")


        
        # Thesis Progress Defense
        if total_no_of_table > 12:
            table_data = self.tables_with_titles[12]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[12] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[12]}")




        file_path = os.path.join(self.output_dir, f"{new_file}.xlsx")
        cell_mappings = {
            1: ['G9', 'G12'],
            2: ['G14'],
            3: ['G17'],
            4: ['G25'],
            5: ['G23', 'G24'],
            6: ['G27'],
            7: ['G18'],
            8: ['G29'],
            9: ['G16'],
            10: ['G20'],
            11: ['G28'],
            12: ['G26']
        }

        for i in range(1, total_no_of_table):
            if os.path.exists(file_path):
                try:
                    if matching_values[i] != 0:
                        print("Inserting data at ", f"{new_file}.xlsx")
                        wb = xw.Book(file_path)
                        sheet = wb.sheets.active
                        cell_locations = cell_mappings.get(i, [])
                        for cell in cell_locations:
                            sheet.range(cell).value = matching_values[i]
                        wb.save(file_path)
                        wb.close()
                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")
            else:
                print(f"File {file_path} does not exist.")




        print(matching_values)
        



    def process_first_excel(self):
        if self.output_dir and self.tables_with_titles and self.sample_excel:
            file_count = 0  # Counter for the files being created
            combined_df = pd.DataFrame()  # Initialize an empty DataFrame to hold all tables

            for i, data in enumerate(self.tables_with_titles):
                table = data["Table"]
                df = pd.DataFrame(table)
                df.ffill(axis=0, inplace=True)
                if i == 0:  # Working with the first table
                    first_table_df_name = df.iloc[:, 1]  # Extracting the content from the second column
                    first_table_df_designation = df.iloc[:, 2]  # Extracting the content from the second column

                    # Create separate Excel files based on each row's content
                    for row_i, (name, designation) in enumerate(zip(first_table_df_name,first_table_df_designation)):
                        if row_i != 0 and row_i != len(first_table_df_name) - 1:
                            name=name.split(',')[0]
                            designation=designation.split(',')[0]
                            new_file=name.replace(" ", "").replace(".", "").replace(",", "")
                            self.new_files.append(new_file)  # Append new_file to the global array
                            new_file_name = new_file + ".xlsx"
                            print(f"Creating {new_file_name}...")
                            shutil.copy(self.sample_excel, os.path.join(self.output_dir, new_file_name))
                            self.print_matching_value_for_file(new_file,name,designation)
                            file_count += 1  # Increment file count

                    # Append the first table content to the combined DataFrame
                    combined_df = pd.concat([combined_df, first_table_df_name], axis=1)
                else:
                    combined_df = pd.concat([combined_df, df.iloc[1:-1, 1]], axis=1)
            print("The total no of files are :",file_count)
            print("The files are: ",self.new_files)
            messagebox.showinfo("Congratulations!", f"Excel Created Successfully! Total Files Created: {file_count}")
            self.clear_labels()
        else:
            messagebox.showwarning("No Tables Found or No Sample Excel", "No tables were detected in the Word document or no Sample Excel selected.")




    def main(self):
        root = tk.Tk()
        root.title("Word to Excel Converter")
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
        style.map("TButton", background=[("active", "#0056b3")])
        main_frame = tk.Frame(root, bg="#f0f0f0")
        main_frame.pack(padx=20, pady=20)

        select_button = ttk.Button(main_frame, text="Select Word Doc", command=self.select_docx, style="TButton")
        select_button.pack(pady=10)
        self.docx_label = tk.Label(main_frame, text="Selected Word Doc: ", bg="#f0f0f0")
        self.docx_label.pack()

        select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=self.select_sample_excel, style="TButton")
        select_sample_button.pack(pady=10)
        self.sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
        self.sample_label.pack()

        generate_button = ttk.Button(main_frame, text="Generate Table in Excel", command=self.generate_excel_from_docx, style="TButton")
        generate_button.pack(pady=10)

        process_button = ttk.Button(main_frame, text="Process the first table", command=self.process_first_excel, style="TButton")
        process_button.pack(pady=10)



        root.mainloop()


if __name__ == "__main__":
    converter = WordToExcelConverter()
    converter.main()
```
```bash
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
from docx import Document
from openpyxl import load_workbook
import xlwings as xw
import re
from docx import Document
from num2words import num2words
from googletrans import Translator
import time



class WordToExcelConverter:

    
    def __init__(self):
        self.docx_file = None
        self.sample_excel = None
        self.output_dir = None
        self.tables_with_titles = None
        self.combined_excel_path = None
        self.new_files = []  # Array to store new_file values globally
        # self.paused = False
        self.total_no_of_table=0
        self.year=0
        self.term=0
        self.dept="CSE"

        
    
    def clear_labels(self):
        self.docx_label.config(text="Selected Word Doc: ")
        self.docx_label.pack()
        self.sample_label.config(text="Selected Sample Excel: ")
        self.sample_label.pack()

    def select_docx(self):
        # Function to handle selection of Word document
        self.docx_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if self.docx_file:
            self.docx_label.config(text=f"Selected Word Doc: {self.docx_file}")
            self.docx_label.pack()

    def select_sample_excel(self):
        # Function to handle selection of Sample Excel file
        self.sample_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.sample_excel:
            self.sample_label.config(text=f"Selected Sample Excel: {self.sample_excel}")
            self.sample_label.pack()


    # def pause_execution(self):
    #     while self.paused:
    #         time.sleep(1)

    # def toggle_pause(self):
    #     self.paused = not self.paused
    #     if self.paused:
    #         self.pause_button.config(state="disabled")
    #         self.continue_button.config(state="active")
    #     else:
    #         self.pause_button.config(state="active")
    #         self.continue_button.config(state="disabled")

    
    def extract_words_before_table(self):
        pattern = r'Bills - (\w+).*?year (\w+)'
        doc = Document(self.docx_file)
        for paragraph in doc.paragraphs:
            for table in doc.tables:
                if paragraph in table._element.iterancestors('w:tbl'):
                    return None, None  # Stop searching when a paragraph is part of a table
                
            match = re.search(pattern, paragraph.text)
            if match:
                self.year = match.group(1)
                self.term = match.group(2)
                # print("Year, Term: ",self.year, self.term)
                print("Year & Term extracted Successfully")
                # messagebox.showinfo("Year & Term extracted Successfully", f"Year & Term extracted Successfully!")
                return

        print("Year & Term extraction failed")
        # messagebox.showwarning("No Year & Term Found", "No Year & Term Found!")

    def extract_data_from_docx(self):
        # Function to extract data from a Word document

        try:
            self.extract_words_before_table()
    
            if self.year and self.term:
                print("Word after 'bills :", self.year)
                print("Word after 'year':", self.term)
            else:
                print("No match found before the table.")
        except FileNotFoundError:
            print(f"Error: The file '{self.docx_file}' was not found.")
        except Exception as e:
            print(f"Error: {e}")
        
        doc = Document(self.docx_file)
        text_content = ""
        self.tables_with_titles = []
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                table_data.append(row_data)
            # Extracting the title before the table
            title = ""
            for paragraph in table.rows[0].cells[0].paragraphs:
                title += paragraph.text
            self.tables_with_titles.append({"Title": title, "Table": table_data})
            self.total_no_of_table = len(self.tables_with_titles)
        return text_content, self.tables_with_titles

    def generate_excel_from_docx(self):
        if self.docx_file:
            text_content, self.tables_with_titles = self.extract_data_from_docx()
            self.output_dir = filedialog.askdirectory()
            temp_folder = os.path.join(self.output_dir, "AllTables")  # Path to AllTables directory
            if self.output_dir:
                if not os.path.exists(temp_folder):
                    os.makedirs(temp_folder)  # Create AllTables directory if it doesn't exist
                
                if self.tables_with_titles:
                    excel_path = os.path.join(temp_folder, "all_tables.xlsx")  # Save all_tables.xlsx inside AllTables
                    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                        for i, data in enumerate(self.tables_with_titles):
                            table = data["Table"]
                            df = pd.DataFrame(table)
                            df.ffill(axis=0, inplace=True)
                            sheet_name = f"Table_{i}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"{sheet_name} added to Excel")
                        messagebox.showinfo("Excel Created Successfully", f"All tables moved to {excel_path}!")
                        # self.clear_labels()
                    # self.pause_execution()
                else:
                    messagebox.showwarning("No Tables Found", "No tables were detected in the Word document.")
            else:
                messagebox.showwarning("Debug", "No output directory selected.")
        else:
            messagebox.showwarning("Oops!", "Please select a valid doc file.")


    def english_to_bengali_number_in_words(self, english_number):
        # Convert English number to words using Indian numbering system
        words_in_english = num2words(english_number, lang='en_IN')
        # Translate to Bengali
        translator = Translator()
        words_in_bengali = translator.translate(words_in_english, dest='bn').text
        # Remove commas and add "টাকা মাত্র" at the end
        modified_output = "কথায় : " + words_in_bengali.replace(',', '') + " টাকা মাত্র।"
        return modified_output


    def print_matching_value_for_file(self, new_file, name, designation):
        print("Processing...")
        total_no_of_table = len(self.tables_with_titles) #12

        # An array of the size of the first, initially all value is 0
        matching_values = [0] * total_no_of_table #12

        # Set Name, Year, Term
        print("Name: ",name)
        print("Designation: ",designation)
        print("Year, Term: ",self.year,self.term)


        # Question Paper Setter & Script Examiner 
        if total_no_of_table > 1:
            table_data = self.tables_with_titles[1]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[1] += float(table_df.iloc[row_idx, 3]) 
                    print(f"Matching value for {new_file}: {matching_values[1]}")

        # Examiners of Class Tests
        if total_no_of_table > 2:
            table_data = self.tables_with_titles[2]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[2] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[2]}")

        # Examiners of Sessional Classes
        if total_no_of_table > 3:
            table_data = self.tables_with_titles[3]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[3] += float(table_df.iloc[row_idx, 2])*float(table_df.iloc[row_idx, 3])/1.5
                    print(f"Matching value for {new_file}: {matching_values[3]}")

        # Script Scrutinizer
        if total_no_of_table > 4:
            table_data = self.tables_with_titles[4]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 0]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[4] += float(table_df.iloc[row_idx, 1]) 
                    print(f"Matching value for {new_file}: {matching_values[4]}")


        # Tabulation & Verification
        if total_no_of_table > 5:
            table_data = self.tables_with_titles[5]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[5] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[5]}")


        # Typing and Drawing
        if total_no_of_table > 6:
            table_data = self.tables_with_titles[6]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 0]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[6] += float(table_df.iloc[row_idx, 1]) 
                    print(f"Matching value for {new_file}: {matching_values[6]}")


        # Central Viva-Voce
        if total_no_of_table > 7:
            table_data = self.tables_with_titles[7]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[7] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[7]}")


        # Student Advising
        if total_no_of_table > 8:
            table_data = self.tables_with_titles[8]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[8] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[8]}")


        # Seminar (CSE 4120) 1 + 1 =2
        if total_no_of_table > 9:
            table_data = self.tables_with_titles[9]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[9] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[9]}")


        # Thesis Progress Defense
        if total_no_of_table > 10:
            table_data = self.tables_with_titles[10]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[10] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[10]}")


        # Final Grade Sheet Verification
        if total_no_of_table > 11:
            table_data = self.tables_with_titles[11]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    print(table_df)
                    matching_values[11] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[11]}")


        
        # Thesis Progress Defense
        if total_no_of_table > 12:
            table_data = self.tables_with_titles[12]["Table"]
            table_df = pd.DataFrame(table_data)

            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                if str(new_file).lower() in table_value.lower() or table_value.lower() in str(new_file).lower():
                    matching_values[12] += float(table_df.iloc[row_idx, 2]) 
                    print(f"Matching value for {new_file}: {matching_values[12]}")




        file_path = os.path.join(self.output_dir, f"{new_file}.xlsx")
        cell_mappings = {
            1: ['G9', 'G12'],
            2: ['G14'],
            3: ['G17'],
            4: ['G25'],
            5: ['G23', 'G24'],
            6: ['G27'],
            7: ['G18'],
            8: ['G29'],
            9: ['G16'],
            10: ['G20'],
            11: ['G28'],
            12: ['G26']
        }

        for i in range(1, total_no_of_table):
            if os.path.exists(file_path):
                try:
                    if matching_values[i] != 0:
                        print("Inserting data at ", f"{new_file}.xlsx")
                        wb = xw.Book(file_path)
                        sheet = wb.sheets.active
                        cell_locations = cell_mappings.get(i, [])
                        for cell in cell_locations:
                            sheet.range(cell).value = matching_values[i]
                        wb.save(file_path)
                        wb.close()
                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")
            else:
                print(f"File {file_path} does not exist.")


        # Open the workbook
        wb = xw.Book(file_path)
        sheet = wb.sheets.active
        # Read value from cell I31
        amount_str = sheet.range('I32').value
        english_str = str(amount_str).split('.')[0]    #type casting the float into string and taking the integer portion only
        english_number= int(english_str)
        bengali_words = self.english_to_bengali_number_in_words(english_number)
        wb.save(file_path)
        wb.close()
        print(bengali_words)

        print(matching_values)

        



    def process_first_excel(self):
        if self.output_dir and self.tables_with_titles and self.sample_excel:
            file_count = 0  # Counter for the files being created
            combined_df = pd.DataFrame()  # Initialize an empty DataFrame to hold all tables

            for i, data in enumerate(self.tables_with_titles):
                table = data["Table"]
                df = pd.DataFrame(table)
                df.ffill(axis=0, inplace=True)
                if i == 0:  # Working with the first table
                    first_table_df_name = df.iloc[:, 1]  # Extracting the content from the second column
                    first_table_df_designation = df.iloc[:, 2]  # Extracting the content from the second column

                    # Create separate Excel files based on each row's content
                    for row_i, (name, designation) in enumerate(zip(first_table_df_name,first_table_df_designation)):
                        if row_i != 0 and row_i != len(first_table_df_name) - 1:
                            name=name.split(',')[0]
                            designation=designation.split(',')[0]
                            new_file=name.replace(" ", "").replace(".", "").replace(",", "")
                            self.new_files.append(new_file)  # Append new_file to the global array
                            new_file_name = new_file + ".xlsx"
                            print(f"Creating {new_file_name}...")
                            shutil.copy(self.sample_excel, os.path.join(self.output_dir, new_file_name))
                            self.print_matching_value_for_file(new_file,name,designation)
                            file_count += 1  # Increment file count


                    # Append the first table content to the combined DataFrame
                    combined_df = pd.concat([combined_df, first_table_df_name], axis=1)
                else:
                    combined_df = pd.concat([combined_df, df.iloc[1:-1, 1]], axis=1)

            print("The total no of files are:", file_count)
            print("The files are:", self.new_files)
            messagebox.showinfo("Congratulations!", f"Excel Created Successfully! Total Files Created: {file_count}")
            self.clear_labels()
        else:
            messagebox.showwarning("No Tables Found or No Sample Excel", "No tables were detected in the Word document or no Sample Excel selected.")




    def main(self):
        root = tk.Tk()
        root.title("KUET teachers' automatic bill generator")
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", foreground="black", background="green")
        style.map("TButton", background=[("active", "#0056b3")])
        main_frame = tk.Frame(root, bg="#f0f0f0")
        main_frame.pack(padx=20, pady=20)

        select_button = ttk.Button(main_frame, text="Select Word Doc", command=self.select_docx, style="TButton")
        select_button.pack(pady=10)

        self.docx_label = tk.Label(main_frame, text="Selected Word Doc: ", bg="#f0f0f0")
        self.docx_label.pack()

        select_sample_button = ttk.Button(main_frame, text="Select Sample Excel", command=self.select_sample_excel, style="TButton")
        select_sample_button.pack(pady=10)

        self.sample_label = tk.Label(main_frame, text="Selected Sample Excel: ", bg="#f0f0f0")
        self.sample_label.pack()

        generate_button = ttk.Button(main_frame, text="Generate Table in Excel", command=self.generate_excel_from_docx, style="TButton")
        generate_button.pack(pady=10)

        process_button = ttk.Button(main_frame, text="Process the first table", command=self.process_first_excel, style="TButton")
        process_button.pack(pady=10)

        
        # pause_button = ttk.Button(main_frame, text="Pause", command=self.toggle_pause, style="TButton")
        # pause_button.pack(pady=10)

        # continue_button = ttk.Button(main_frame, text="Continue", command=self.toggle_pause, style="TButton")
        # continue_button.pack(pady=10)
        # continue_button.configure(state="disabled")


        root.mainloop()

if __name__ == "__main__":
    converter = WordToExcelConverter()
    converter.main()
```



## Reference
Visit these sites for more info. 

https://pypdf2.readthedocs.io/en/3.0.0/

https://pypi.org/project/pdfplumber/

https://openpyxl.readthedocs.io/en/stable/index.html

https://pandas.pydata.org/

https://matplotlib.org/