import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd # type: ignore

def on_focus_in(event):
    """Change the text color to green on focus in."""
    event.widget.config(fg="black")

def on_focus_out(event):
    """Reset the text color to black on focus out if the widget is empty."""
    if not event.widget.get():
        event.widget.config(fg="black")

def update_column_mapping(RSRP_columns, CINR_columns):
    """Update column mapping based on user input."""
    column_mapping = {}
    
    for i, col in enumerate(RSRP_columns):
        if col:
            column_mapping[col] = f'R{i+1} CINR (0)'
    
    for i, col in enumerate(CINR_columns):
        if col:
            column_mapping[f'R{i+1} RSRP (dBm)'] = col
    
    return column_mapping

def process_csv_files(directory, RSRP_columns, CINR_columns, RSRP_conditions, CINR_conditions):
    """Process CSV files in a directory."""
    results = []
    if not os.path.isdir(directory):
        messagebox.showerror("Error", "Invalid directory selected.")
        return results
    
    for entry in os.scandir(directory):
        if entry.is_dir():
            results.extend(process_csv_files(entry.path, RSRP_columns, CINR_columns, RSRP_conditions, CINR_conditions))
        elif entry.is_file() and entry.name.endswith('.csv'):
            try:
                df = pd.read_csv(entry.path)
                column_mapping = update_column_mapping(RSRP_columns, CINR_columns)
                
                if any(col in df.columns for col in column_mapping.keys()):
                    RSRP_cols = [col for col in RSRP_columns if col]
                    CINR_cols = [col for col in CINR_columns if col]
                    
                    df_filtered_RSRP = df[df[RSRP_cols].apply(lambda x: x >= RSRP_conditions).any(axis=1)]
                    df_filtered_CINR = df[df[CINR_cols].apply(lambda x: x >= CINR_conditions).any(axis=1)]
                    
                    percentage_RSRP = (len(df_filtered_RSRP) / len(df)) * 100
                    percentage_CINR = (len(df_filtered_CINR) / len(df)) * 100
                    
                    results.append({'File': entry.name, 'RSRP': percentage_RSRP, 'CINR': percentage_CINR})
            except Exception as e:
                messagebox.showwarning("Warning", f"Error processing {entry.name}: {str(e)}")
    
    return results

def run_analysis():
    """Run the analysis based on user inputs."""
    directory = filedialog.askdirectory(title="Select Directory")
    if directory:
        try:
            RSRP_columns = [RSRP1_entry.get(), RSRP2_entry.get(), RSRP3_entry.get()]
            CINR_columns = [CINR1_entry.get(), CINR2_entry.get(), CINR3_entry.get()]
            
            # Handle conditions
            RSRP_conditions = None
            CINR_conditions = None
            
            if RSRP_condition_entry.get():
                try:
                    RSRP_conditions = int(RSRP_condition_entry.get())
                except ValueError:
                    messagebox.showwarning("Warning", "RSRP Condition should be an integer.")
                    return
            
            if CINR_condition_entry.get():
                try:
                    CINR_conditions = int(CINR_condition_entry.get())
                except ValueError:
                    messagebox.showwarning("Warning", "CINR Condition should be an integer.")
                    return
            
            results = process_csv_files(directory, RSRP_columns, CINR_columns, RSRP_conditions, CINR_conditions)
            output_df = pd.DataFrame(results)
            output_file = os.path.join(directory, "Output.xlsx")
            output_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Analysis completed. Results saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main application window
root = tk.Tk()
root.title("DTS_Software")

root.configure(background="#000")

# Main section frame
section_frame = tk.Frame(root, bg="#000")
section_frame.pack(fill=tk.BOTH, expand=True)

# Animated gradient background

# Create a frame for layout
frame = tk.Frame(root, padx=120, pady=100, bg="#e1d0ba", highlightbackground="#000")
root.title("DTS")
frame.pack()

signin_frame = tk.Frame(section_frame, bg="#000", bd=0)
signin_frame.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

# Content frame inside signin_frame
content_frame = tk.Frame(signin_frame, bg="#000")
content_frame.pack(pady=40)

# Header
header_label = tk.Label(content_frame, text="DTS", font=("Chiller",180, "bold"), bg="#000", fg="#e1d0ba")
header_label.pack(pady=20)

# Labels and Entries for RSRP columns and conditions
RSRP_label = tk.Label(frame, text="RSRP Columns:", bg="#e1d0ba", font=('Helvetica Bold', 12))
RSRP_label.grid(row=0, column=0, padx=10, pady=5)

RSRP1_entry = tk.Entry(frame, bg="lightgray")
RSRP1_entry.grid(row=1, column=1, padx=5, pady=5)
RSRP1_entry.bind("<FocusIn>", on_focus_in)
RSRP1_entry.bind("<FocusOut>", on_focus_out)

RSRP2_entry = tk.Entry(frame, bg="lightgray")
RSRP2_entry.grid(row=2, column=1, padx=5, pady=5)

RSRP2_entry.bind("<FocusIn>", on_focus_in)
RSRP2_entry.bind("<FocusOut>", on_focus_out)

RSRP3_entry = tk.Entry(frame, bg="lightgray")
RSRP3_entry.grid(row=3, column=1, padx=5, pady=5)
RSRP3_entry.bind("<FocusIn>", on_focus_in)
RSRP3_entry.bind("<FocusOut>", on_focus_out)

RSRP_condition_label = tk.Label(frame, text="RSRP Condition:",bg="#e1d0ba",font=('Helvetica Bold', 12))
RSRP_condition_label.grid(row=7, column=0, padx=10, pady=5)


RSRP_condition_entry = tk.Entry(frame, bg="lightgray")
RSRP_condition_entry.grid(row=7, column=1, padx=10, pady=5)

# Labels and Entries for CINR columns and conditions
CINR_label = tk.Label(frame, text="CINR Columns:",bg="#e1d0ba",font=('Helvetica Bold', 12))
CINR_label.grid(row=0, column=2, padx=10, pady=5)

CINR1_entry = tk.Entry(frame, bg="lightgray")
CINR1_entry.grid(row=1, column=3, padx=5, pady=5)
CINR1_entry.bind("<FocusIn>", on_focus_in)
CINR1_entry.bind("<FocusOut>", on_focus_out)

CINR2_entry = tk.Entry(frame, bg="lightgray")
CINR2_entry.grid(row=2, column=3, padx=5, pady=5)
CINR2_entry.bind("<FocusIn>", on_focus_in)
CINR2_entry.bind("<FocusOut>", on_focus_out)

CINR3_entry = tk.Entry(frame, bg="lightgray")
CINR3_entry.grid(row=3, column=3, padx=5, pady=5)
CINR3_entry.bind("<FocusIn>", on_focus_in)
CINR3_entry.bind("<FocusOut>", on_focus_out)

CINR_condition_label = tk.Label(frame, text="CINR Condition:",bg="#e1d0ba",font=('Helvetica Bold', 12))
CINR_condition_label.grid(row=7, column=2, padx=10, pady=5)

CINR_condition_entry = tk.Entry(frame, bg="lightgray")
CINR_condition_entry.grid(row=7, column=3, padx=10, pady=5)

# Create a button to run the analysis


def exit_application():
    """Exit the application."""
    root.destroy()
    
def on_hover(event, button):
    """Change button appearance on hover."""
    button.config(bg="lightgreen", fg="black")

def on_leave(event, button):
    """Restore button appearance when mouse leaves."""
    button.config(bg="#e1d0ba", fg="#000")

# Run Button
run_button = tk.Button(frame, text="Run Analysis", command=run_analysis,
                       bg="#e1d0ba", fg="#000",
                       borderwidth=2, relief="groove",
                       padx=10, pady=5)
run_button.grid(row=10, column=1, columnspan=2, padx=10, pady=10)
run_button.bind("<Enter>", lambda event, btn=run_button: on_hover(event, btn))
run_button.bind("<Leave>", lambda event, btn=run_button: on_leave(event, btn))

def refresh_interface():
    """Refresh the interface."""
    # Clear all the entry fields
    RSRP1_entry.delete(0, tk.END)
    RSRP2_entry.delete(0, tk.END)
    RSRP3_entry.delete(0, tk.END)
    
    CINR1_entry.delete(0, tk.END)
    CINR2_entry.delete(0, tk.END)
    CINR3_entry.delete(0, tk.END)
    
    RSRP_condition_entry.delete(0, tk.END)
    CINR_condition_entry.delete(0, tk.END)

# Refresh Button
refresh_button = tk.Button(frame, text="Refresh", command=refresh_interface,
                           bg="#e1d0ba", fg="#000",
                           borderwidth=2, relief="groove",
                           padx=10, pady=5)
refresh_button.grid(row=10, column=0, columnspan=1, padx=10, pady=10)
refresh_button.bind("<Enter>", lambda event, btn=refresh_button: on_hover(event, btn))
refresh_button.bind("<Leave>", lambda event, btn=refresh_button: on_leave(event, btn))

# Exit Button
exit_button = tk.Button(frame, text="Exit", command=exit_application,
                        bg="#e1d0ba", fg="#000",
                        borderwidth=2, relief="groove",
                        padx=10, pady=5)
exit_button.grid(row=10, column=3, columnspan=4, padx=10, pady=10)
exit_button.bind("<Enter>", lambda event, btn=exit_button: on_hover(event, btn))
exit_button.bind("<Leave>", lambda event, btn=exit_button: on_leave(event, btn))

# Run the Tkinter event loop
root.mainloop()
