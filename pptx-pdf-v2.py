import os
import comtypes.client
from pptxtopdf import convert
import ctypes
import sys

if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # Fallback for older Windows
        except:
            pass

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def select_folder(title = "Select Output Folder"):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title = title)
    root.destroy()
    return folder_path

def select_pptx_files(title="Select PowerPoint Files"):   
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    file_types = [
        ("PowerPoint files", "*.pptx *.ppt"),
        ("All files", "*.*")
    ]
    
    file_paths = filedialog.askopenfilenames(
        title=title,
        filetypes=file_types
    )
    root.destroy()
    return list(file_paths)

input_dir = select_pptx_files("Select PowerPoint file to convert")
if not input_dir:
    print("No file selected. Exiting.")
    exit()

output_dir = select_folder("Select Output directory (where PDF will be saved)")
if not output_dir:
    print("No folder selected. Exiting.")
    exit()
    
for i in range(len(input_dir)):
    convert(input_dir[i], output_dir)
    
print("Conversion completed.")