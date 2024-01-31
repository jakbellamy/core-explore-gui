import tkinter as tk
from tkinter import ttk
import os
import pandas as pd
import re

def get_sequences():
    """
    This function is used to get all unique 6-digit sequences from the file names.
    """
    files = os.listdir('./processed_cores')
    sequences = set()
    for file in files:
        matches = re.findall(r'\d{6}', file)
        for match in matches:
            sequences.add(match)
    return sequences

def get_files(sequence=None):
    """
    This function is used to get the files in the ./processed_cores directory for the user to select from.
    The names of the files will be displayed as clean, but the actual file names will be returned.
    If a sequence is provided, only files containing that sequence will be returned.
    """
    files = os.listdir('./processed_cores')
    if sequence:
        files = [file for file in files if sequence in file]
    return {show_name_as_clean(f): f for f in files}

def on_sequence_select(event):
    """
    Function called when a sequence is selected in the dropdown.
    """
    selected_sequence = sequence_var.get()
    global file_mapping
    file_mapping = get_files(selected_sequence)
    file_var.set('')
    file_dropdown['values'] = list(file_mapping.keys())

def get_data(file):
    """
    This function is used to get the data from the selected file.
    """
    data = pd.read_csv(f'./processed_cores/{file}')
    return data

def show_name_as_clean(file):
    """
    This function is used to show the file name as clean to the user.
    """
    return file.split('.')[0]

def on_run():
    """
    Function called when the Run button is pressed.
    """
    selected_file_clean = file_var.get()
    selected_file_raw = file_mapping[selected_file_clean]
    data = get_data(selected_file_raw)
    # Copy data to clipboard using Pandas
    data.to_clipboard(index=False, excel=True)
    print(f"Copied data from {selected_file_raw} to clipboard")

# Set up the main application window
root = tk.Tk()
root.title("Core Explore Kit")

# Sequence Selection
tk.Label(root, text="Select Date:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
sequence_var = tk.StringVar()
sequence_dropdown = ttk.Combobox(root, textvariable=sequence_var)
sequence_dropdown['values'] = sorted(list(get_sequences()), reverse=True)
sequence_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky='we')
sequence_dropdown.bind('<<ComboboxSelected>>', on_sequence_select)

# File Selection
tk.Label(root, text="Select File:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
file_mapping = get_files()
file_var = tk.StringVar()
file_dropdown = ttk.Combobox(root, textvariable=file_var)
file_dropdown['values'] = sorted(list(get_files()))
file_dropdown.grid(row=1, column=1, padx=10, pady=5, sticky='we')

# Run Button
run_button = ttk.Button(root, text="Run", command=on_run)
run_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='we')

# Configure column 1 to expand with window size
root.columnconfigure(1, weight=1)

# Start the Tkinter event loop
root.mainloop()
