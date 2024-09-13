import tkinter as tk
from tkinter import filedialog

def ask_for_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    file_path = filedialog.askopenfilename(
        title="Select Microsoft Project File", filetypes=[("Microsoft Project files", "*.mpp")]
    )
    return file_path
