import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from openpyxl import Workbook
import logging
from .core import process_file
from .config import DEBUG_SAVE_RAW

loaded_files = []
master_workbook = None

def load_files():
    global loaded_files
    file_paths = filedialog.askopenfilenames(
        title="Select PDF or Excel files",
        filetypes=[("PDF and Excel Files", "*.pdf *.xls *.xlsx"), ("All Files", "*.*")]
    )
    if file_paths:
        loaded_files = file_paths
        messagebox.showinfo("Files Loaded", f"Loaded {len(file_paths)} files.")

def load_master_workbook():
    global master_workbook
    file_path = filedialog.asksaveasfilename(
        title="Select or Create Master Workbook",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if file_path:
        master_workbook = file_path
        messagebox.showinfo("Master Workbook", f"Using master workbook:\n{file_path}")

def run_extraction():
    if not loaded_files:
        messagebox.showerror("Error", "No files loaded.")
        return
    if not master_workbook:
        messagebox.showerror("Error", "No master workbook set.")
        return
    create_new = False
    if not os.path.exists(master_workbook):
        create_new = True
    else:
        try:
            _ = pd.ExcelFile(master_workbook)
        except Exception:
            messagebox.showwarning(
                "Corrupt Workbook",
                f"The file at {master_workbook} is not a valid Excel file.\nA new master workbook will be created."
            )
            os.remove(master_workbook)
            create_new = True
    if create_new:
        wb = Workbook()
        ws = wb.active
        ws.title = "Summary"
        wb.save(master_workbook)
    with pd.ExcelWriter(master_workbook, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        for f in loaded_files:
            logging.info(f"Processing {f}")
            process_file(f, writer)
    messagebox.showinfo("Done", f"Processed {len(loaded_files)} file(s).\nSaved to:\n{master_workbook}")

def main():
    root = tk.Tk()
    root.title("Company File Processor")
    root.geometry("500x250")
    label = tk.Label(root, text="Load files and master workbook, then run extraction.", font=("Arial", 12))
    label.pack(pady=15)
    btn_load_files = tk.Button(root, text="Load Files", command=load_files, width=25, height=2)
    btn_load_files.pack(pady=5)
    btn_load_master = tk.Button(root, text="Load Master Workbook", command=load_master_workbook, width=25, height=2)
    btn_load_master.pack(pady=5)
    btn_run = tk.Button(root, text="Run Extraction", command=run_extraction, width=25, height=2)
    btn_run.pack(pady=15)
    root.mainloop()

if __name__ == "__main__":
    main()
