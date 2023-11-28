import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
from docx import Document


class WordToExcelConverter:
    def __init__(self):
        self.docx_file = None
        self.sample_excel = None
        self.output_dir = None
        self.tables_with_titles = None
        self.combined_excel_path = None
    
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
                        for idx, data in enumerate(self.tables_with_titles):
                            table = data["Table"]
                            df = pd.DataFrame(table)
                            df.ffill(axis=0, inplace=True)
                            sheet_name = f"Table_{idx}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"{sheet_name} added to Excel")
                        messagebox.showinfo("Excel Created Successfully", f"All tables moved to {excel_path}!")
                        self.clear_labels()
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

            for idx, data in enumerate(self.tables_with_titles):
                table = data["Table"]
                df = pd.DataFrame(table)
                df.ffill(axis=0, inplace=True)
                if idx == 0:  # Working with the first table
                    first_table_df = df.iloc[:, 1]  # Extracting the content from the second column

                    # Create separate Excel files based on each row's content
                    for row_idx, value in enumerate(first_table_df):
                        if row_idx != 0 and row_idx != len(first_table_df) - 1:
                            new_file_name = value.replace(" ", "").replace(".", "").replace(",", "") + ".xlsx"
                            print(f"Creating {new_file_name}...")
                            shutil.copy(self.sample_excel, os.path.join(self.output_dir, new_file_name))
                            file_count += 1  # Increment file count

                    # Append the first table content to the combined DataFrame
                    combined_df = pd.concat([combined_df, first_table_df], axis=1)
                else:
                    combined_df = pd.concat([combined_df, df.iloc[1:-1, 1]], axis=1)
            print(file_count)
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
