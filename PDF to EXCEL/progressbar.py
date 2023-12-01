import tkinter as tk
from tkinter import ttk

class ProgressBarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Progress Bar Example")

        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress.pack(pady=20)

        self.start_button = ttk.Button(self.root, text="Start", command=self.start_progress)
        self.start_button.pack(pady=10)

    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()  # Refresh the window to update the progress bar

    def start_progress(self):
        # Simulating a process with multiple stages
        total_stages = 5
        for i in range(total_stages):
            # Perform some operation for each stage
            # Here, update the progress bar
            self.update_progress((i + 1) * (100 / total_stages))
            self.root.after(500)  # Simulate some delay between stages (500 milliseconds)

if __name__ == "__main__":
    root = tk.Tk()
    app = ProgressBarApp(root)
    root.mainloop()
