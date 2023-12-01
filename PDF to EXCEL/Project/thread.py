import tkinter as tk
import threading
import time

class CounterThread(threading.Thread):
    def __init__(self, gui):
        threading.Thread.__init__(self)
        self.gui = gui
        self.running = False
        self.count = 0

    def run(self):
        self.running = True
        while self.running:
            self.count += 1
            self.gui.update_counter_label(self.count)
            time.sleep(1)

    def stop(self):
        self.running = False

class GUI:
    def __init__(self, root):
        self.root = root
        self.counter_thread = None

        self.counter_label = tk.Label(root, text="Counter: 0")
        self.counter_label.pack()

        self.start_button = tk.Button(root, text="Start", command=self.start_counter)
        self.start_button.pack()

        self.stop_button = tk.Button(root, text="Stop", command=self.stop_counter)
        self.stop_button.pack()

    def start_counter(self):
        if not self.counter_thread or not self.counter_thread.is_alive():
            self.counter_thread = CounterThread(self)
            self.counter_thread.start()

    def stop_counter(self):
        if self.counter_thread:
            self.counter_thread.stop()

    def update_counter_label(self, count):
        self.counter_label.config(text=f"Counter: {count}")

def main():
    root = tk.Tk()
    root.title("Threaded Counter")

    gui = GUI(root)
    
    root.mainloop()

if __name__ == "__main__":
    main()
