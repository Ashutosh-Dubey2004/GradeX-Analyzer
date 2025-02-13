import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import os
import datetime

# Function placeholders (to be implemented)
def fetch_results():
    messagebox.showinfo("Fetch", "Fetching results...")

def view_sheet():
    messagebox.showinfo("View", "Viewing sheet...")

def delete_sheet():
    messagebox.showinfo("Delete", "Deleting sheet...")

def download_sheet():
    messagebox.showinfo("Download", "Downloading sheet...")

def show_developer_info():
    dev_window = tk.Toplevel(root)
    dev_window.title("Developer Info")
    dev_window.geometry("600x250")
    dev_window.resizable(False, False)
    dev_window.configure(bg="#f0f0f0")
    FALLBACK_INFO = """
Developed by:

* Ashutosh Dubey\t\t(IMCA 2021 batch)
\t\t\t\t
Student at Acropolis FCA department under guidance of \t\t
Prof. Nitin Kulkarni.

For help and assistance email at: 
ashutoshdubey.ca21@acropolis.in
"""

    dev_label = tk.Label(dev_window, text=FALLBACK_INFO, font=("Arial", 12, "bold"), bg="#f0f0f0", fg="black",  justify='left')
    dev_label.pack(pady=15)

# Main window setup
root = tk.Tk()
root.title("RGPV AutoFetch - Automated Result Fetcher")
root.geometry("1100x650")
root.resizable(False, False)
root.configure(bg="#f0f0f0")

# Title label
title_label = tk.Label(root, text="RGPV AutoFetch", font=("Arial", 16, "bold"), bg="#003366", fg="white", padx=20, pady=10)
title_label.pack(fill=tk.X)

# Horizontal Separator (Below Title)
title_separator = ttk.Separator(root, orient="horizontal")
title_separator.pack(fill="x")


# Footer section
footer_frame = tk.Frame(root, bg="#003366")
footer_frame.pack(side=tk.BOTTOM, fill=tk.X)

footer_label = tk.Label(footer_frame, text="Made by student at FCA department", font=("Arial", 10), bg="#003366", fg="white", pady=5)
footer_label.pack(side=tk.LEFT, padx=10)

# Small Button on Right Side of footer
info_button = tk.Button(footer_frame, text="ℹ️", command=show_developer_info, font=("Arial", 10, "bold"), bg="#ffffff", fg="black", width=3)
info_button.pack(side=tk.RIGHT, padx=10, pady=2)

# Left Panel
left_frame = tk.Frame(root, padx=15, pady=15, relief=tk.GROOVE, borderwidth=2, bg="#d9e6f2")
left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

# Label for Left Panel (Title)
left_title = tk.Label(left_frame, text="Input Section", font=("Arial", 12, "bold"), bg="#d9e6f2", fg="#003366")
left_title.pack(anchor="w", padx=5, pady=5)  

# Vertical Separator (Between Left & Center Panel)
left_separator = ttk.Separator(root, orient="vertical")
left_separator.pack(side="left", fill="y")

# Course Dropdown
course_label = tk.Label(left_frame, text="Course:", font=("Arial", 12), bg="#d9e6f2")
course_label.pack(anchor='w', pady=5)
courses = ["DDMCA", "MCA"]
course_var = tk.StringVar()
course_dropdown = ttk.Combobox(left_frame, values=courses, textvariable=course_var, state="readonly", width=22)
course_dropdown.pack(pady=5)

# Semester Dropdown
sem_label = tk.Label(left_frame, text="Semester:", font=("Arial", 12), bg="#d9e6f2")
sem_label.pack(anchor='w', pady=5)
semesters = [str(i) for i in range(1, 11)]
sem_var = tk.StringVar()
sem_dropdown = ttk.Combobox(left_frame, values=semesters, textvariable=sem_var, state="readonly", width=22)
sem_dropdown.pack(pady=5)

# Date Picker
date_label = tk.Label(left_frame, text="Date:", font=("Arial", 12), bg="#d9e6f2")
date_label.pack(anchor='w', pady=5)
date_picker = DateEntry(left_frame, date_pattern='yyyy-mm-dd', width=20)
date_picker.pack(pady=12)

# Roll No Prefix
prefix_label = tk.Label(left_frame, text="Roll No. Prefix:", font=("Arial", 12), bg="#d9e6f2")
prefix_label.pack(anchor='w', pady=5)
prefix_entry = tk.Entry(left_frame, width=24)
prefix_entry.pack(pady=5)

# Roll No Start
start_label = tk.Label(left_frame, text="Roll No. Start:", font=("Arial", 12), bg="#d9e6f2")
start_label.pack(anchor='w', pady=5)
start_entry = tk.Entry(left_frame, width=24)
start_entry.pack(pady=5)

# Roll No End
end_label = tk.Label(left_frame, text="Roll No. End:", font=("Arial", 12), bg="#d9e6f2")
end_label.pack(anchor='w', pady=5)
end_entry = tk.Entry(left_frame, width=24)
end_entry.pack(pady=5)

# Fetch Button
fetch_button = tk.Button(left_frame, text="Fetch Results", command=fetch_results, font=("Arial", 12), width=22, bg="#003366", fg="white")
fetch_button.pack(pady=20)

# Center Panel (Sheet List)
center_frame = tk.Frame(root, padx=15, pady=15, relief=tk.GROOVE, borderwidth=2, bg="#ffffff")
center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Add a title label above the listbox
sheet_title = tk.Label(center_frame, text="Generated Sheets", font=("Arial", 12, "bold"), bg="#ffffff", fg="#003366")
sheet_title.pack(anchor="w", padx=10, pady=5)

# Listbox for displaying sheets
sheet_listbox = tk.Listbox(center_frame, font=("Arial", 12))
sheet_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# Right Panel
right_frame = tk.Frame(root, padx=15, pady=15, relief=tk.GROOVE, borderwidth=2, bg="#d9e6f2")
right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)

# Label for Right Panel (Title)
right_title = tk.Label(right_frame, text="Report Actions", font=("Arial", 12, "bold"), bg="#d9e6f2", fg="#003366")
right_title.pack(anchor="w", padx=5, pady=5) 

# Vertical Separator (Between Center & Right Panel)
right_separator = ttk.Separator(root, orient="vertical")
right_separator.pack(side="right", fill="y")


filter_label = tk.Label(right_frame, text="Filter:", font=("Arial", 12), bg="#d9e6f2")
filter_label.pack(anchor='w', pady=7)
filter_entry = tk.Entry(right_frame, width=24)
filter_entry.pack(pady=7)

view_button = tk.Button(right_frame, text="View Sheet", command=view_sheet, font=("Arial", 12), width=22, bg="#003366", fg="white")
view_button.pack(pady=15)

delete_button = tk.Button(right_frame, text="Delete Sheet", command=delete_sheet, font=("Arial", 12), width=22, bg="#cc0000", fg="white")
delete_button.pack(pady=15)

download_button = tk.Button(right_frame, text="Download Sheet", command=download_sheet, font=("Arial", 12), width=22, bg="#009900", fg="white")
download_button.pack(pady=15)




root.mainloop()