import tkinter as tk
from tkinter import ttk

class ProgressBarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Progress Bar Example")

        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.start_button = ttk.Button(self.root, text="Start", command=self.start_progress)
        self.reset_button = ttk.Button(self.root, text="Reset", command=self.reset_progress)

        # Initially hide the progress bar
        self.progress.pack_forget()

        self.start_button.pack(pady=10)
        self.reset_button.pack(pady=10)

    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()

    def start_progress(self):
        self.progress.pack(pady=20)  # Display the progress bar
        # Simulating a process with multiple stages
        total_stages = 5
        for i in range(total_stages):
            self.update_progress((i + 1) * (100 / total_stages))
            self.root.after(500)

    def reset_progress(self):
        self.progress.pack_forget()  # Hide the progress bar
        self.progress['value'] = 0   # Reset the progress bar to the initial position

if __name__ == "__main__":
    root = tk.Tk()
    app = ProgressBarApp(root)
    root.mainloop()
