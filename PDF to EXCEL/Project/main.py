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
        self.dept="সিএসই"    # Mother Department
        self.translator = Translator()
        self.dept_suffixes_mapping = {
             
        "computer science and engineering": "সিএসই",
        "computer science & engineering": "সিএসই",
        "electrical and electronic engineering": "ইইই",
        "electrical & electronic engineering": "ইইই",
        "electronics and communication engineering": "ইসিই",
        "electronics & communication engineering": "ইসিই",
        "biomedical engineering": "বিএমই",
        "materials science and engineering": "এমএসই",
        "materials science & engineering": "এমএসই",
        "civil engineering": "পুরকৌশল",
        "urban and regional planning": "ইউআরপি",
        "urban & regional planning": "ইউআরপি",
        "building engineering and construction management": "বিইসিএম",
        "building engineering & construction management": "বিইসিএম",
        "architecture": "স্থাপত্য",
        "mathematics": "গণিত",
        "math": "গণিত",
        "chemistry": "রসায়ন",
        "physics": "পদার্থ",
        "humanities": "মানবিক",
        "mechanical engineering": "যন্ত্র প্রকৌশল",
        "industrial engineering and management": "শিল্প প্রকৌশল",
        "industrial engineering & management": "শিল্প প্রকৌশল",
        "energy science and engineering": "ইএসই",
        "energy science & engineering": "ইএসই",
        "leather engineering": "লেদার",
        "textile engineering": "টেক্সটাইল",
        "chemical engineering": "টেক্সটাইল",
        "mechatronics engineering": "মেকাট্রনিক্স",
        }

    def dept_translate_to_bengali(self, english_text):
        bengali_text = self.dept_suffixes_mapping.get(english_text.lower())
        if not bengali_text:
            # If the translation is not found in the mapping, use Google Translate
            translated = self.translator.translate(english_text, dest='bn')
            bengali_text = translated.text
        return bengali_text
        
    
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


    def extract_department_line(self):
        pattern = r'(?:Department of|Department Of)(.*)'
        doc = Document(self.docx_file)
        for paragraph in doc.paragraphs:
            match = re.search(pattern, paragraph.text)
            if match:
                self.dept = match.group(1).strip()  # Extract text after the department pattern
                return

        return None

    
    def extract_year_and_term(self):
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
            self.extract_department_line()

            if self.dept:
                print("Department Line:", self.dept)
            else:
                print("No 'Department of' line found.")
        except FileNotFoundError:
            print(f"Error: The file '{self.docs_file}' was not found.")
        except Exception as e:
            print(f"Error: {e}")

        try:
            self.extract_year_and_term()
    
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

    def should_skip_translation(self, text):
        name_patterns = [r'Dean', r'Md\.', r'Dr\.', r'Sk\.', r'Fatema']
        for pattern in name_patterns:
            if re.search(pattern, text):
                return True
        return False

    def translate_to_bengali(self, text):
        translator = Translator()
    
        # Define the translation rules
        translation_rules = {
            r'Dean': 'ডিন',
            r'Md\.': 'মোঃ',
            r'Dr\.': 'ড.',
            r'Sk\.': 'শেখ',
            r'Most': 'মোসাম্মৎ',
            r'Fatema': 'ফাতেমা'
        }

        parts = text.split()
        translated_parts = []
        for part in parts:
            if not self.should_skip_translation(part):
                # Apply the specific translation rule if found
                for pattern, replacement in translation_rules.items():
                    if re.search(pattern, part):
                        part = re.sub(pattern, replacement, part)
                        break
                translated_part = translator.translate(part, dest='bn').text
            else:
                # Use provided translation rules when skipping translation
                for pattern, replacement in translation_rules.items():
                    if re.search(pattern, part):
                        translated_part = re.sub(pattern, replacement, part)
                        break
            translated_parts.append(translated_part)

        return ' '.join(translated_parts)


    def english_to_bengali_number_in_words(self, english_number):
        # Convert English number to words using Indian numbering system
        words_in_english = num2words(english_number, lang='en_IN')
        # Translate to Bengali
        translator = Translator()
        words_in_bengali = translator.translate(words_in_english, dest='bn').text
        # Remove commas and add "টাকা মাত্র" at the end
        modified_output = "কথায় : " + words_in_bengali.replace(',', '') + " টাকা মাত্র।"
        return modified_output
    
    def convert_suffix_to_bengali(self, text):
        # Dictionary mapping English suffixes to Bengali
        suffixes_mapping = {
            "1st": "১ম",
            "2nd": "২য়",
            "3rd": "৩য়",
            "4th": "৪র্থ",  # You can add more mappings as needed
            # Add more mappings for other suffixes
        }

        # Replace English suffixes with Bengali equivalents
        for suffix in suffixes_mapping:
            if suffix in text:
                text = text.replace(suffix, suffixes_mapping[suffix])

        return text

    def convert_year_term_suffixes_to_bengali(self, text):
        # Dictionary mapping English year_term_suffixes to Bengali
        year_term_suffixes_mapping = {
            "1st": "১ম",
            "2nd": "২য়",
            "3rd": "৩য়",
            "4th": "৪র্থ",  # You can add more mappings as needed
            # Add more mappings for other year_term_suffixes
        }

        # Replace English year_term_suffixes with Bengali equivalents
        for suffix in year_term_suffixes_mapping:
            if suffix in text:
                text = text.replace(suffix, year_term_suffixes_mapping[suffix])

        return text


    def print_matching_value_for_file(self, new_file, name, designation, department):
        print("Processing...")
        total_no_of_table = len(self.tables_with_titles) #12

        # An array of the size of the first, initially all value is 0
        matching_values = [0] * total_no_of_table #12



        # Set Name, Year, Term
        print("Name: ",name)
        # name = self.translate_to_bengali(name)
        # print("Name: ",name)

        print("Designation: ",designation)
        designation = self.translate_to_bengali(designation)
        print("Designation: ",designation)

        print("Dept: ",self.dept)
        self.dept = self.dept_translate_to_bengali(self.dept.lower())
        print("Dept: ",self.dept)

        print("Department: ",department)
        department = self.dept_translate_to_bengali(department.lower())
        print("Department: ",department)
        
        print("Year: ",self.year)
        print("Term: ",self.term)
        self.year=self.convert_year_term_suffixes_to_bengali(self.year)
        self.term=self.convert_year_term_suffixes_to_bengali(self.term)
        print("Year: ",self.year)
        print("Term: ",self.term)




        # Question Paper Setter & Script Examiner 
        if total_no_of_table > 1:
            table_data = self.tables_with_titles[1]["Table"]
            table_df = pd.DataFrame(table_data)
            print("Lets see: ")
            for row_idx in range(1, len(table_df)):
                table_value = str(table_df.iloc[row_idx, 1]).replace(" ", "").replace(".", "").replace(",", "")
                print(table_value)
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
                    for row_i, (name, designation_and_department) in enumerate(zip(first_table_df_name,first_table_df_designation)):
                        if row_i != 0 and row_i != len(first_table_df_name) - 1:
                            name=name.split(',')[0]
                            designation=designation_and_department.split(',')[0]
                            department=designation_and_department.split(',')[1]
                            new_file=name.replace(" ", "").replace(".", "").replace(",", "")
                            self.new_files.append(new_file)  # Append new_file to the global array
                            new_file_name = new_file + ".xlsx"
                            print(f"Creating {new_file_name}...")
                            shutil.copy(self.sample_excel, os.path.join(self.output_dir, new_file_name))
                            self.print_matching_value_for_file(new_file,name,designation,department)
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