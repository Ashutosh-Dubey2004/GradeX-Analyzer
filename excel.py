import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import os
import pandas as pd
import subprocess

DBLOC = f"{os.environ['USERPROFILE']}\\.RGPV AutoFetch\\.result-data\\"
os.makedirs(DBLOC, exist_ok=True)

def refresh_sheet_list(sheet_listbox):
    sheet_listbox.delete(0, tk.END)
    result_files = os.listdir(DBLOC)
    for file in result_files:
        if file.endswith(".xlsx"):
            sheet_listbox.insert(tk.END, file)

def save_to_excel(data, course, semester, year, sheet_listbox):
    filename = f"{course}_Sem{semester}_{year}.xlsx"
    filepath = os.path.join(DBLOC, filename)
    df = pd.DataFrame(data[1:], columns=data[0])
    df.to_excel(filepath, index=False)
    refresh_sheet_list(sheet_listbox)
    

def download(sheet_listbox):
    selected = sheet_listbox.curselection()
    if not selected:
        messagebox.showerror("Error", "No sheet selected!")
        return
    filename = sheet_listbox.get(selected[0])
    src_path = os.path.join(DBLOC, filename)
    dest_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")],initialfile=filename)
    if dest_path:
        os.replace(src_path, dest_path)
        messagebox.showinfo("Success", "Sheet downloaded successfully!")
        refresh_sheet_list(sheet_listbox)

def delete(sheet_listbox):
    selected = sheet_listbox.curselection()
    if not selected:
        messagebox.showerror("Error", "No sheet selected!")
        return
    filename = sheet_listbox.get(selected[0])
    filepath = os.path.join(DBLOC, filename)
    if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {filename}?"):
        os.remove(filepath)
        refresh_sheet_list(sheet_listbox)
        messagebox.showinfo("Deleted", "Sheet deleted successfully!")

def view(sheet_listbox):
    selected = sheet_listbox.curselection()
    if not selected:
        messagebox.showerror("Error", "No sheet selected!")
        return
    filename = sheet_listbox.get(selected[0])
    filepath = os.path.join(DBLOC, filename)
    subprocess.run(["start", "", filepath], shell=True)

def filter(event, sheet_listbox, filter_entry):
    """Filters the listbox based on user input (Course, Semester, Branch)."""
    query = filter_entry.get().strip().lower()
    
    if not query:  
        refresh_sheet_list(sheet_listbox)  # Show all files if query is empty
        return  

    sheet_listbox.delete(0, tk.END)
    result_files = os.listdir(DBLOC)

    for file in result_files:
        if file.endswith(".xlsx"):
            parts = file.replace(".xlsx", "").split("_")  
            if len(parts) >= 3:
                course_name, sem_part, branch_name = parts[0].lower(), parts[1].lower(), parts[2].lower()

                course_match = query == course_name  
                sem_match = query in sem_part
                branch_match = query in branch_name

                # Show file only if it matches Course, Semester, or Branch
                if course_match or sem_match or branch_match:
                    sheet_listbox.insert(tk.END, file)