import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry

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