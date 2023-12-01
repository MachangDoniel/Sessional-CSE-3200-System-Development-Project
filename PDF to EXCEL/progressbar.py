import tkinter as tk
from tkinter import messagebox
import threading
import time
from tkinter.ttk import Progressbar  # Import Progressbar from ttk

class ProgressBarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Progress Bar Example")

        self.progress = tk.IntVar()
        self.progress_bar = Progressbar(self.root, orient=tk.HORIZONTAL, length=200, mode='determinate', variable=self.progress)
        self.start_button = tk.Button(self.root, text="Start", command=self.start_progress)
        self.pause_button = tk.Button(self.root, text="Pause", command=self.pause_progress)
        self.continue_button = tk.Button(self.root, text="Continue", command=self.continue_progress)
        self.reset_button = tk.Button(self.root, text="Reset", command=self.reset_progress)

        self.progress_bar.pack(pady=20)
        self.start_button.pack(pady=10)
        self.pause_button.pack(pady=10)
        self.continue_button.pack(pady=10)
        self.reset_button.pack(pady=10)

        self.paused = False
        self.completed = False
        self.pause_condition = threading.Condition()
        self.progress.set(0)

    def update_progress(self):
        while self.progress.get() < 100:
            if not self.paused:
                self.progress.set(self.progress.get() + 1)
            time.sleep(0.1)

        self.completed = True
        messagebox.showinfo("Completed", "Process Completed!")

    def start_progress(self):
        self.paused = False
        self.completed = False
        self.progress.set(0)
        self.thread = threading.Thread(target=self.update_progress)
        self.thread.start()

    def pause_progress(self):
        self.paused = True

    def continue_progress(self):
        with self.pause_condition:
            self.paused = False
            self.pause_condition.notify()

    def reset_progress(self):
        self.progress.set(0)
        self.paused = False
        if hasattr(self, 'thread'):
            self.thread.join()

if __name__ == "__main__":
    root = tk.Tk()
    app = ProgressBarApp(root)
    root.mainloop()
