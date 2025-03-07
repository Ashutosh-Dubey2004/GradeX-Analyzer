import tkinter as tk
from tkinter import ttk, messagebox

import optimized
from optimized import retrieveMultipleResults
import excel
import threading
import winsound  # For playing sound

def download_sheet():
    excel.download(sheet_listbox)

def delete_sheet():
    excel.delete(sheet_listbox)

def view_sheet():
    excel.view(sheet_listbox)

# Function to validate inputs
def validate_inputs():
    course = course_var.get().upper()
    sem = sem_var.get().strip()
    prefixRollNO = prefix_entry.get().strip()
    startRollNo = start_entry.get().strip()
    endRollNo = end_entry.get().strip()
    batch = batch_entry.get().strip()

    # Course validation
    if course not in ["DDMCA", "MCA"]:
        messagebox.showerror("Input Error", "Invalid Course! Select DDMCA or MCA.")
        return None

    # Semester validation
    if not sem.isdigit() or int(sem) not in range(1, 11):
        messagebox.showerror("Input Error", "Semester must be between 1 and 10.")
        return None

    # Batch validation (only 2 digits allowed)
    if not batch.isdigit() or len(batch) != 2:
        messagebox.showerror("Input Error", "Batch must be exactly 2 digits.")
        return None

    # Roll No. Prefix validation (exactly 10 characters)
    if len(prefixRollNO) != 10:
        messagebox.showerror("Input Error", "Enter a valid Roll No. Prefix (exactly 10 characters).")
        return None

    # Start and End Roll Number validation
    if not startRollNo.isdigit() or not endRollNo.isdigit():
        messagebox.showerror("Input Error", "Roll Numbers must be numeric values.")
        return None

    startRollNo, endRollNo = int(startRollNo), int(endRollNo)
    
    if startRollNo > endRollNo:
        messagebox.showerror("Input Error", "Start Roll No. cannot be greater than End Roll No.")
        return None
    
    # If all validations pass, return the values
    return course, sem, prefixRollNO, startRollNo, endRollNo, batch

# Fetch Results
loop_thread = None
fetching_window = None  # Global reference to store popup window
optimized.abort = False  # Variable to track cancellation

def fetch():
    global loop_thread
    optimized.abort = False  # Reset abort flag before starting
    loop_thread = threading.Thread(target=fetch_result, daemon=True)
    loop_thread.start()

def fetch_result():
    validated_data = validate_inputs()
    if validated_data:
        course, sem, prefixRollNO, startRollNo, endRollNo, batch = validated_data
        
        show_fetching_popup()  # Show fetching popup before starting
        
        if optimized.abort:  # If cancelled, stop execution
            close_fetching_popup()
            return  

        data = retrieveMultipleResults(course, sem, prefixRollNO, startRollNo, endRollNo)

        if optimized.abort:  # Double-check before processing data
            close_fetching_popup()
            return  

        if data:
            excel.save_to_excel(data, course, sem, batch, sheet_listbox)  # Save fetched data
            root.after(0, lambda: (close_fetching_popup(), messagebox.showinfo("Success", "Results fetched and saved successfully!")))
        else:
            root.after(0, lambda: (close_fetching_popup(),messagebox.showerror("Error", "No results fetched. Please try again!")))
    else:
        root.after(0, lambda: messagebox.showerror("Error", "Invalid inputs. Please check and try again!"))

def show_fetching_popup():
    """Creates Fetching window"""
    global fetching_window  
    fetching_window = tk.Toplevel(root)
    fetching_window.title("Fetching...")
    fetching_window.geometry("380x170")  
    fetching_window.configure(bg="#f8f9fa")  
    fetching_window.resizable(False, False)

    winsound.MessageBeep(winsound.MB_ICONASTERISK)  # Play Windows notification sound

    # Disable interaction with main window
    fetching_window.transient(root)  
    fetching_window.grab_set() 

    frame = tk.Frame(fetching_window, bg="white", relief="groove", borderwidth=2)
    frame.pack(expand=True, fill="both", padx=15, pady=15)

    label = tk.Label(frame, text="Fetching results, please wait...", 
                     font=("Segoe UI", 11, "bold"), bg="white", fg="black", anchor="w")
    label.pack(side="top", pady=(5, 0), padx=10, anchor="w")

    # Progress indicator (loading animation)
    progress = ttk.Progressbar(frame, mode="indeterminate", length=250)
    progress.pack(pady=(10, 10))
    progress.start(10)  
    
    cancel_button = tk.Button(frame, text="Cancel Fetching", command=cancel_fetching, font=("Arial", 12), width=22, bg="#cc0000", fg="white")
    cancel_button.pack(pady=15)

    center_window(fetching_window)

def close_fetching_popup():
    """Closes fetching popup if it's open"""
    global fetching_window
    if fetching_window:
        fetching_window.destroy()
        fetching_window = None  # Reset reference

def cancel_fetching():
    """Handles cancellation of fetching process"""
    global loop_thread
    optimized.abort = True  # Set abort flag

    if loop_thread and loop_thread.is_alive():
        loop_thread.join(timeout=1)  # Ensure thread stops execution
    
    close_fetching_popup()  # Close fetching popup
    messagebox.showerror("Cancelled", "Fetching process was cancelled!")  

def stop():
    """Stops the fetching process"""
    optimized.abort = True 

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

def update_prefix():
    course = course_var.get().upper()
    batch = batch_entry.get().strip()
    
    if batch.isdigit() and len(batch) == 2:  # Ensure batch is numeric and 2 digits
        if course == "DDMCA":
            prefix_entry.delete(0, tk.END)
            prefix_entry.insert(0, f"0827CA{batch}DD")
        elif course == "MCA":
            prefix_entry.delete(0, tk.END)
            prefix_entry.insert(0, f"0827CA{batch}10")

def center_window(window):
    """Centers any given Tkinter window on the screen."""
    window.update_idletasks()
    x = (window.winfo_screenwidth() - window.winfo_width()) // 2
    y = (window.winfo_screenheight() - window.winfo_height()) // 2
    window.geometry(f"+{x}+{y}")

# Main window setup
root = tk.Tk()
root.title("GradeX Analyzer – Efficient Result Extraction, Smarter Data Use")
root.geometry("1100x650")
root.resizable(False, False)
root.configure(bg="#f0f0f0")
center_window(root)

# Title label
title_label = tk.Label(root, text="GradeX Analyzer", font=("Arial", 16, "bold"), bg="#003366", fg="white", padx=20, pady=10)
title_label.pack(fill=tk.X)

# Footer section
footer_frame = tk.Frame(root, bg="#003366")
footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
footer_label = tk.Label(footer_frame, text="Made by student at FCA department", font=("Arial", 10), bg="#003366", fg="white", pady=5)
footer_label.pack(side=tk.LEFT, padx=10)
info_button = tk.Button(footer_frame, text="ℹ️", command=show_developer_info, font=("Arial", 10, "bold"), bg="#ffffff", fg="black", width=3)
info_button.pack(side=tk.RIGHT, padx=10, pady=2)

# Left Panel
left_frame = tk.Frame(root, padx=15, pady=15, relief=tk.GROOVE, borderwidth=2, bg="#d9e6f2")
left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
left_title = tk.Label(left_frame, text="Input Section", font=("Arial", 12, "bold"), bg="#d9e6f2", fg="#003366")
left_title.pack(anchor="w", padx=5, pady=5)

# Course Dropdown
course_label = tk.Label(left_frame, text="Course:", font=("Arial", 12), bg="#d9e6f2")
course_label.pack(anchor='w', pady=5)
courses = ["DDMCA", "MCA"]
course_var = tk.StringVar()
course_dropdown = ttk.Combobox(left_frame, values=courses, textvariable=course_var, state="readonly", width=22)
course_dropdown.pack(pady=5)
course_dropdown.bind("<<ComboboxSelected>>", lambda event: update_prefix())# Course selection change

# Semester Dropdown
sem_label = tk.Label(left_frame, text="Semester:", font=("Arial", 12), bg="#d9e6f2")
sem_label.pack(anchor='w', pady=5)
semesters = [str(i) for i in range(1, 11)]
sem_var = tk.StringVar()
sem_dropdown = ttk.Combobox(left_frame, values=semesters, textvariable=sem_var, state="readonly", width=22)
sem_dropdown.pack(pady=5)

# Date Picker
batch_label = tk.Label(left_frame, text="Batch:", font=("Arial", 12), bg="#d9e6f2")
batch_label.pack(anchor='w', pady=5)
batch_entry = tk.Entry(left_frame, width=24)
batch_entry.pack(pady=5)
batch_entry.bind("<KeyRelease>", lambda event: update_prefix())# Batch entry change

# Roll No Inputs
prefix_label = tk.Label(left_frame, text="Roll No. Prefix:", font=("Arial", 12), bg="#d9e6f2")
prefix_label.pack(anchor='w', pady=5)
prefix_entry = tk.Entry(left_frame, width=24)
prefix_entry.pack(pady=5)

start_label = tk.Label(left_frame, text="Roll No. Start:", font=("Arial", 12), bg="#d9e6f2")
start_label.pack(anchor='w', pady=5)
start_entry = tk.Entry(left_frame, width=24)
start_entry.pack(pady=5)

end_label = tk.Label(left_frame, text="Roll No. End:", font=("Arial", 12), bg="#d9e6f2")
end_label.pack(anchor='w', pady=5)
end_entry = tk.Entry(left_frame, width=24)
end_entry.pack(pady=5)

# Fetch Button
fetch_button = tk.Button(left_frame, text="Fetch Results", command=fetch, font=("Arial", 12), width=22, bg="#003366", fg="white")
fetch_button.pack(pady=20)

# Center Panel (Sheet List)
center_frame = tk.Frame(root, padx=15, pady=15, relief=tk.GROOVE, borderwidth=2, bg="#ffffff")
center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Add a title label above the listbox
sheet_title = tk.Label(center_frame, text="Generated Sheets", font=("Arial", 12, "bold"), bg="#ffffff", fg="#003366")
sheet_title.pack(anchor="w", padx=10, pady=5)

# Listbox for displaying sheets
sheet_listbox = tk.Listbox(
    center_frame, 
    font=("Segoe UI", 12),  # For Font
    bd=2,  # Border
    highlightthickness=0,  
    relief=tk.FLAT,  # look smoother
    selectbackground="#cce5ff",  # Light blue selection
    selectforeground="black",  # Black text on selection
)
sheet_listbox.pack(fill=tk.BOTH, expand=True)

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
filter_entry.bind("<KeyRelease>", lambda event: excel.filter(event, sheet_listbox, filter_entry))

view_button = tk.Button(right_frame, text="View Sheet", command=view_sheet, font=("Arial", 12), width=22, bg="#003366", fg="white")
view_button.pack(pady=15)

delete_button = tk.Button(right_frame, text="Delete Sheet", command=delete_sheet, font=("Arial", 12), width=22, bg="#cc0000", fg="white")
delete_button.pack(pady=15)

download_button = tk.Button(right_frame, text="Download Sheet", command=download_sheet, font=("Arial", 12), width=22, bg="#009900", fg="white")
download_button.pack(pady=15)

excel.refresh_sheet_list(sheet_listbox)
root.mainloop()
