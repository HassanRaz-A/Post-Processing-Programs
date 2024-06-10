import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd

def on_focus_in(event):
    """Change the text color to black on focus in."""
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
                    RSRP_cols = [col for col in RSRP_columns if col in df.columns]
                    CINR_cols = [col for col in CINR_columns if col in df.columns]

                    # Drop rows with missing values in the specified columns
                    df = df.dropna(subset=RSRP_cols + CINR_cols)

                    df_filtered_RSRP = df[df[RSRP_cols].apply(lambda x: x >= RSRP_conditions).any(axis=1)]
                    df_filtered_CINR = df[df[CINR_cols].apply(lambda x: x >= CINR_conditions).any(axis=1)]

                    percentage_RSRP = round((len(df_filtered_RSRP) / len(df)) * 100) if len(df) > 0 else 0
                    percentage_CINR = round((len(df_filtered_CINR) / len(df)) * 100) if len(df) > 0 else 0

                    results.append({'File': entry.name, 'RSRP': f"{percentage_RSRP}%", 'CINR': f"{percentage_CINR}%"})
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

def exit_application():
    """Exit the application."""
    root.destroy()
    
def on_hover(event, button):
    """Change button appearance on hover."""
    button.config(bg="lightgreen", fg="black")

def on_leave(event, button):
    """Restore button appearance when mouse leaves."""
    button.config(bg="#e1d0ba", fg="#000")

def refresh_interface():
    """Refresh the interface."""
    RSRP1_entry.set('')
    RSRP2_entry.set('')
    RSRP3_entry.set('')
    CINR1_entry.set('')
    CINR2_entry.set('')
    CINR3_entry.set('')
    RSRP_condition_entry.set('')
    CINR_condition_entry.set('')

# Predefined column names and conditions
predefined_rsrp_columns = ['R0 RSRP (0)', 'R0 RSRP (1)', 'SSB_RP (0)', 'SSB_RP (1)']
predefined_cinr_columns = ['R0 RS CINR (0)', 'SSB_CINR (0)']
predefined_conditions_rsrp = [str(i) for i in range(-120, -60, 5)]  # Example conditions for RSRP
predefined_conditions_cinr = [str(i) for i in range(1, 11)]  # Predefined conditions for CINR (1 to 10)

# Create the main application window
root = tk.Tk()
root.title("DTS_Software")
root.configure(background="#000")

# Create a frame for layout
frame = tk.Frame(root, padx=120, pady=100, bg="#e1d0ba", highlightbackground="#000")
frame.pack()

# Labels and Comboboxes for RSRP columns and conditions
RSRP_label = tk.Label(frame, text="RSRP Columns:", bg="#e1d0ba", font=('Helvetica Bold', 12))
RSRP_label.grid(row=0, column=0, padx=10, pady=5)

RSRP1_entry = ttk.Combobox(frame, values=predefined_rsrp_columns)
RSRP1_entry.grid(row=1, column=1, padx=5, pady=5)

RSRP2_entry = ttk.Combobox(frame, values=predefined_rsrp_columns)
RSRP2_entry.grid(row=2, column=1, padx=5, pady=5)

RSRP3_entry = ttk.Combobox(frame, values=predefined_rsrp_columns)
RSRP3_entry.grid(row=3, column=1, padx=5, pady=5)

RSRP_condition_label = tk.Label(frame, text="RSRP Condition:", bg="#e1d0ba", font=('Helvetica Bold', 12))
RSRP_condition_label.grid(row=7, column=0, padx=10, pady=5)

RSRP_condition_entry = ttk.Combobox(frame, values=predefined_conditions_rsrp)
RSRP_condition_entry.grid(row=7, column=1, padx=10, pady=5)

# Labels and Comboboxes for CINR columns and conditions
CINR_label = tk.Label(frame, text="CINR Columns:", bg="#e1d0ba", font=('Helvetica Bold', 12))
CINR_label.grid(row=0, column=2, padx=10, pady=5)

CINR1_entry = ttk.Combobox(frame, values=predefined_cinr_columns)
CINR1_entry.grid(row=1, column=3, padx=5, pady=5)

CINR2_entry = ttk.Combobox(frame, values=predefined_cinr_columns)
CINR2_entry.grid(row=2, column=3, padx=5, pady=5)

CINR3_entry = ttk.Combobox(frame, values=predefined_cinr_columns)
CINR3_entry.grid(row=3, column=3, padx=5, pady=5)

CINR_condition_label = tk.Label(frame, text="CINR Condition:", bg="#e1d0ba", font=('Helvetica Bold', 12))
CINR_condition_label.grid(row=7, column=2, padx=10, pady=5)

CINR_condition_entry = ttk.Combobox(frame, values=predefined_conditions_cinr)
CINR_condition_entry.grid(row=7, column=3, padx=10, pady=5)

# Run Button
run_button = tk.Button(frame, text="Run Analysis", command=run_analysis,
                       bg="#e1d0ba", fg="#000",
                       borderwidth=2, relief="groove",
                       padx=10, pady=5)
run_button.grid(row=10, column=1, columnspan=2, padx=10, pady=10)
run_button.bind("<Enter>", lambda event, btn=run_button: on_hover(event, btn))
run_button.bind("<Leave>", lambda event, btn=run_button: on_leave(event, btn))

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

