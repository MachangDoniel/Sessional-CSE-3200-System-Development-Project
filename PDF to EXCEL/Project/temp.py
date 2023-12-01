import tkinter as tk
from tkinter import ttk

class AutomaticBillGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatic bill generator")

        main_frame = tk.Frame(root, bg="#f0f0f0")
        main_frame.pack(padx=20, pady=20)

        # Title Label
        title_label = tk.Label(main_frame, text="Automatic bill generator", font=("Arial", 16), pady=10)
        title_label.pack()

        # Divide the window horizontally
        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill=tk.BOTH, expand=True)

        left_frame = tk.Frame(top_frame)
        left_frame.pack(side=tk.LEFT, padx=10, pady=10)

        right_frame = tk.Frame(top_frame, bg="white")  # Empty frame for future content
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Buttons on the left side
        select_button = ttk.Button(left_frame, text="Select Word Doc")
        select_button.pack(pady=10)

        select_sample_button = ttk.Button(left_frame, text="Select Sample Excel")
        select_sample_button.pack(pady=10)

        generate_button = ttk.Button(left_frame, text="Generate Table in Excel")
        generate_button.pack(pady=10)

        process_button = ttk.Button(left_frame, text="Process the first table")
        process_button.pack(pady=10)

        # Placeholders for future content on the right side
        placeholder_label = tk.Label(right_frame, text="Right side content goes here", font=("Arial", 12))
        placeholder_label.pack(padx=50, pady=50)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomaticBillGeneratorGUI(root)
    root.mainloop()
