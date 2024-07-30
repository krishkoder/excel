import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys

# Increase the recursion limit
sys.setrecursionlimit(5000)

def remove_duplicates(file_path):
    try:
        data = pd.read_excel(file_path)
        data_cleaned = data.drop_duplicates()
        return data_cleaned
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return None

def select_files():
    files_selected = filedialog.askopenfilenames(title="Select Excel files", filetypes=[("Excel files", "*.xlsx")])
    return files_selected

def select_directory(prompt):
    return filedialog.askdirectory(title=prompt)

def select_file(prompt, default_filename):
    return filedialog.asksaveasfilename(title=prompt, defaultextension=".xlsx", initialfile=default_filename, filetypes=[("Excel files", "*.xlsx")])

def prompt_overwrite(file_path):
    answer = messagebox.askyesnocancel("File exists", f"The file {file_path} already exists. Do you want to overwrite it?")
    if answer is None:
        return "cancel"
    return "overwrite" if answer else "rename"

def handle_existing_file(file_path):
    action = prompt_overwrite(file_path)
    if action == "overwrite":
        return file_path
    elif action == "rename":
        return select_file("Save As", os.path.basename(file_path))
    return None

def start_cleaning():
    input_files = select_files()
    if not input_files:
        messagebox.showerror("Error", "No files selected")
        return

    output_directory = select_directory("Select Output Directory")
    if not output_directory:
        messagebox.showerror("Error", "Output directory not specified")
        return

    os.makedirs(output_directory, exist_ok=True)
    combined_output_file_single_sheet = None
    combined_output_file_separate_sheets = None

    combine_single_sheet = combine_single_sheet_var.get()
    combine_separate_sheets = combine_separate_sheets_var.get()

    if combine_single_sheet:
        combined_output_file_single_sheet = select_file("Save the combined Excel file (single sheet)", "cleaned_data_combined_single_sheet.xlsx")
        if not combined_output_file_single_sheet:
            messagebox.showerror("Error", "Combined output file not selected for single sheet")
            return

    if combine_separate_sheets:
        combined_output_file_separate_sheets = select_file("Save the combined Excel file (separate sheets)", "cleaned_data_combined_separate_sheets.xlsx")
        if not combined_output_file_separate_sheets:
            messagebox.showerror("Error", "Combined output file not selected for separate sheets")
            return

    num_files = len(input_files)
    progress_bar["maximum"] = num_files

    try:
        if combine_single_sheet:
            combined_data = pd.DataFrame()
            for i, file_path in enumerate(input_files):
                cleaned_data = remove_duplicates(file_path)
                if cleaned_data is not None:
                    output_file = os.path.join(output_directory, os.path.basename(file_path))
                    if os.path.exists(output_file):
                        output_file = handle_existing_file(output_file)
                        if not output_file:
                            continue
                    combined_data = pd.concat([combined_data, cleaned_data], ignore_index=True)
                    cleaned_data.to_excel(output_file, index=False)
                    print(f'Successfully cleaned and saved file: {output_file}')
                else:
                    print(f'Failed to clean file: {file_path}')
                progress_bar["value"] = i + 1
                root.update_idletasks()
            combined_data.to_excel(combined_output_file_single_sheet, index=False)
            print(f"Combined data saved to: {combined_output_file_single_sheet}")

        if combine_separate_sheets:
            with pd.ExcelWriter(combined_output_file_separate_sheets, engine='xlsxwriter') as writer:
                for i, file_path in enumerate(input_files):
                    file_name = os.path.basename(file_path)
                    cleaned_data = remove_duplicates(file_path)
                    if cleaned_data is not None:
                        sheet_name = os.path.splitext(file_name)[0]
                        cleaned_data.to_excel(writer, sheet_name=sheet_name, index=False)
                        output_file = os.path.join(output_directory, file_name)
                        if os.path.exists(output_file):
                            output_file = handle_existing_file(output_file)
                            if not output_file:
                                continue
                        cleaned_data.to_excel(output_file, index=False)
                        print(f'Successfully cleaned and saved file: {output_file}')
                    else:
                        print(f'Failed to clean file: {file_path}')
                    progress_bar["value"] = i + 1
                    root.update_idletasks()
            print(f"Data saved with separate sheets to: {combined_output_file_separate_sheets}")

        if not combine_single_sheet and not combine_separate_sheets:
            for i, file_path in enumerate(input_files):
                cleaned_data = remove_duplicates(file_path)
                if cleaned_data is not None:
                    output_file = os.path.join(output_directory, os.path.basename(file_path))
                    if os.path.exists(output_file):
                        output_file = handle_existing_file(output_file)
                        if not output_file:
                            continue
                    cleaned_data.to_excel(output_file, index=False)
                    print(f'Successfully cleaned and saved file: {output_file}')
                else:
                    print(f'Failed to clean file: {file_path}')
                progress_bar["value"] = i + 1
                root.update_idletasks()

        messagebox.showinfo("Success", "All cleaned data has been processed and saved.")
        status_label.config(text="Process completed")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        status_label.config(text="Process failed")
        print(f"An error occurred: {e}")
    finally:
        progress_bar["value"] = 0

# Create the main window
root = tk.Tk()
root.title("Excel Cleaner")

# Set the size of the window
root.geometry("800x600")
root.minsize(800, 600)

# Input files entry and button
input_files_frame = tk.Frame(root)
input_files_frame.pack(pady=10)
input_files_label = tk.Label(input_files_frame, text="Input Files:", font=("Helvetica", 12))
input_files_label.pack(side=tk.LEFT, padx=5)
input_files_entry = tk.Entry(input_files_frame, width=50, font=("Helvetica", 12))
input_files_entry.pack(side=tk.LEFT, padx=5)
input_files_button = tk.Button(input_files_frame, text="Browse", command=lambda: input_files_entry.insert(0, ";".join(select_files())), font=("Helvetica", 12))
input_files_button.pack(side=tk.LEFT, padx=5)

# Output directory entry and button
output_dir_frame = tk.Frame(root)
output_dir_frame.pack(pady=10)
output_dir_label = tk.Label(output_dir_frame, text="Output Directory:", font=("Helvetica", 12))
output_dir_label.pack(side=tk.LEFT, padx=5)
output_dir_entry = tk.Entry(output_dir_frame, width=50, font=("Helvetica", 12))
output_dir_entry.pack(side=tk.LEFT, padx=5)
output_dir_button = tk.Button(output_dir_frame, text="Browse", command=lambda: output_dir_entry.insert(0, select_directory("Select Output Directory")), font=("Helvetica", 12))
output_dir_button.pack(side=tk.LEFT, padx=5)

# Combine options checkboxes
combine_options_frame = tk.Frame(root)
combine_options_frame.pack(pady=10)
combine_single_sheet_var = tk.BooleanVar()
combine_separate_sheets_var = tk.BooleanVar()
combine_single_sheet_check = tk.Checkbutton(combine_options_frame, text="Combine into single sheet", variable=combine_single_sheet_var, font=("Helvetica", 12))
combine_separate_sheets_check = tk.Checkbutton(combine_options_frame, text="Combine into separate sheets", variable=combine_separate_sheets_var, font=("Helvetica", 12))
combine_single_sheet_check.pack(side=tk.LEFT, padx=5)
combine_separate_sheets_check.pack(side=tk.LEFT, padx=5)

# Start button
start_button = tk.Button(root, text="Start Cleaning Process", command=start_cleaning, font=("Helvetica", 14))
start_button.pack(pady=20)

# Progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress_bar.pack(pady=10)

# Status label
status_label = tk.Label(root, text="", font=("Helvetica", 12))
status_label.pack(pady=10)

# Run the GUI event loop
root.mainloop()
