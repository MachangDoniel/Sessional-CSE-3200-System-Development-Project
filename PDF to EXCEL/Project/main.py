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