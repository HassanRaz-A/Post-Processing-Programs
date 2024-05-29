import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar, Text, Checkbutton

def get_files(source_folder, file_types):
    files = []
    for file in os.listdir(source_folder):
        if file.endswith('.csv') and 'csv' in file_types:
            files.append(file)
        elif file.endswith('.xlsx') and 'excel' in file_types:
            files.append(file)
    return files

def browse_source_folder():
    source_folder = filedialog.askdirectory()
    source_path_entry.delete(0, tk.END)
    source_path_entry.insert(tk.END, source_folder)
    files = get_files(source_folder, [])
    if csv_var.get():
        files += get_files(source_folder, ['csv'])
    else:
        files += get_files(source_folder, ['excel'])
    files_listbox.delete(0, tk.END)
    for file in files:
        files_listbox.insert(tk.END, file)

def browse_target_folder():
    target_folder = filedialog.askdirectory()
    target_path_entry.delete(0, tk.END)
    target_path_entry.insert(tk.END, target_folder)

def submit_form():
    selected_indices = files_listbox.curselection()
    if not selected_indices:
        messagebox.showwarning("Warning", "No files selected.")
        return

    source_folder = source_path_entry.get()
    target_folder = target_path_entry.get()

    if create_folder_var.get():
        folder_name = folder_name_entry.get()
        target_folder = os.path.join(target_folder, folder_name)
        os.makedirs(target_folder, exist_ok=True)

    new_filenames = rename_entry.get("1.0", tk.END).splitlines()

    if len(new_filenames) != len(selected_indices):
        messagebox.showerror("Error", "Number of new filenames does not match number of selected files.")
        return

    try:
        for index, selected_index in enumerate(selected_indices):
            filename = files_listbox.get(selected_index)
            source_file_path = os.path.join(source_folder, filename)
            target_file_path = os.path.join(target_folder, new_filenames[index])
            shutil.copy2(source_file_path, target_file_path)
        messagebox.showinfo("Success", "Files copied successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error copying files: {e}")

# Create main window
root = tk.Tk()
root.title("CSV File Management For Wind Console")
root.configure(background="#e1d0ba")
# Source Folder
source_label = tk.Label(root,bg="#e1d0ba", text="Source Folder:", font=('Helvetica Bold', 12))
source_label.grid(row=1, column=0, sticky="w")
source_path_entry = tk.Entry(root,bg="#e1d0ba", width=150)
source_path_entry.grid(row=1, column=1, padx=1)
source_button = tk.Button(root,bg="#e1d0ba", text="Browse",width=20, command=browse_source_folder, font=('Helvetica Bold', 10))
source_button.grid(row=1, column=2)

# Create New Folder
create_folder_var = tk.BooleanVar()
create_folder_checkbutton = tk.Checkbutton(root,bg="#e1d0ba", text="Create New Folder", variable=create_folder_var, font=('Helvetica Bold', 12))
create_folder_checkbutton.grid(row=2, column=2, sticky="w")

# Folder Name
folder_name_label = tk.Label(root,bg="#e1d0ba", text="Folder Name:", font=('Helvetica Bold', 12))
folder_name_label.grid(row=2, column=0, sticky="w")
folder_name_entry = tk.Entry(root, bg="#e1d0ba",width=150)
folder_name_entry.grid(row=2, column=1, padx=1)

# Target Folder
target_label = tk.Label(root,bg="#e1d0ba", text="Target Folder:", font=('Helvetica Bold', 12))
target_label.grid(row=3, column=0, sticky="w")
target_path_entry = tk.Entry(root,bg="#e1d0ba", width=150)
target_path_entry.grid(row=3, column=1, padx=1)
target_button = tk.Button(root,bg="#e1d0ba", text="Browse",width=20, command=browse_target_folder, font=('Helvetica Bold', 10))
target_button.grid(row=3, column=2)

# File Types
file_types_label = tk.Label(root,bg="#e1d0ba", text="Select File Types:", font=('Helvetica Bold', 12))
file_types_label.grid(row=4, column=2, sticky="")
csv_var = tk.BooleanVar()
csv_checkbox = Checkbutton(root,bg="#e1d0ba", text="CSV", variable=csv_var, font=('Helvetica Bold', 12))
csv_checkbox.grid(row=5, column=2, sticky="nw")
excel_var = tk.BooleanVar()
excel_checkbox = Checkbutton(root,bg="#e1d0ba", text="Excel", variable=excel_var, font=('Helvetica Bold', 12))
excel_checkbox.grid(row=5, column=2, sticky="ne")


# Files Listbox
files_label = tk.Label(root,bg="#e1d0ba", text="Files:", font=('Helvetica Bold', 12))
files_label.grid(row=5, column=0, sticky="w")
files_listbox = Listbox(root,bg="#e1d0ba", width=150, height=30, selectmode=tk.MULTIPLE)
files_listbox.grid(row=5, column=1, padx=1, pady=5, columnspan=1)
scrollbar = Scrollbar(root,bg="#e1d0ba", orient="vertical", command=files_listbox.yview)
scrollbar.grid(row=5, column=2, padx=0, sticky="ns")
files_listbox.config(yscrollcommand=scrollbar.set)

# Rename Text
rename_label = tk.Label(root, bg="#e1d0ba",text="New File Names (per line):", font=('Helvetica Bold', 12))
rename_label.grid(row=6, column=0, sticky="w")
rename_entry = Text(root,bg="#e1d0ba", width=112, height=5)
rename_entry.grid(row=6, column=1, padx=5, pady=5)

# Submit Button
submit_button = tk.Button(root, text="Execute Files",bg="#e1d0ba",width=20, height=2, command=submit_form, font=('Helvetica Bold', 12))
submit_button.grid(row=7, column=1,columnspan=1, pady=10)

root.mainloop()
