import tkinter as tk
from tkinter import ttk, font, filedialog, messagebox
import os
import pandas as pd
import re
from src.deep_dive_starters import get_vendor_ranks, get_agent_ranks, re_agg_file
import shutil

def update_instructions(text):
    """
    This function is used to update the instructions text.
    """
    instructions.config(state='normal')
    instructions.insert(tk.END, text)
    instructions.config(state='disabled')

def clear_instructions():
    # Enable the widget, clear its contents, then disable it
    instructions.config(state='normal')
    instructions.delete('1.0', tk.END)
    instructions.config(state='disabled')

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
        
    files = sorted(files)  # Sort the files in alphabetical order
    
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
    data = data[data['Ven'] == 'M']

    # Sort data by Associate Name then Closed Sales Volume YTD
    data = data.sort_values(by=['Associate Name', 'Closed Sales Volume YTD'], ascending=[True, False])
    return data

def show_name_as_clean(file):
    """
    This function is used to show the file name as clean to the user.
    """
    return file.split('.')[0]

selected_file_raw = None
rankings_button = None

def on_run():
    """
    Function called when the Run button is pressed.
    """
    global selected_file_raw
    global rankings_button

    clear_instructions()
    update_instructions(f"Core Report Copied to Clipboard\n\nPaste the values to A1 of either YTD12 or LY12. Then click the 'Get Vendor Rankings' button.")

    selected_file_clean = file_var.get()
    selected_file_raw = file_mapping[selected_file_clean]
    data = get_data(selected_file_raw)
    # Copy data to clipboard using Pandas
    data.to_clipboard(index=False, excel=True)
    print(f"Copied data from {selected_file_raw} to clipboard")

    # Remove the "Run" button
    run_button.grid_forget()

    # Add the "Get Vendor Rankings" button
    rankings_button = ttk.Button(root, text="Get Vendor Rankings", command=get_vendor_rankings)
    rankings_button.grid(row=6, column=0, columnspan=2, padx=200, pady=5, sticky='we')

def get_vendor_rankings():
    global selected_file_raw
    global rankings_button

    print('Getting vendor rankings...')
    clear_instructions()
    get_vendor_ranks(selected_file_raw)
    update_instructions(f"Vendor rankings have been copied to the clipboard.\n\nPaste the values to D2 if YTD or J2 if Last Year. Then click the 'Get Agent Rankings' button.")

    # Remove the "Get Vendor Rankings" button
    rankings_button.grid_forget()

    # Add the "Get Agent Rankings" button
    rankings_button = ttk.Button(root, text="Get Agent Rankings", command=get_agent_rankings)
    rankings_button.grid(row=6, column=0, columnspan=2, padx=200, pady=5, sticky='we')


def get_agent_rankings():
    global selected_file_raw
    global rankings_button

    print('Getting agent rankings...')
    clear_instructions()
    get_agent_ranks(selected_file_raw)
    update_instructions(f"Agent rankings have been copied to the clipboard.\n\nPaste the values to G2 if YTD or M2 if Last Year. Then click the 'Re-Aggregate' button.")

    # Remove the "Get Agent Rankings" button
    rankings_button.grid_forget()

    # Add the "Re-Aggregate" button
    rankings_button = ttk.Button(root, text="Finish", command=return_to_start)
    rankings_button.grid(row=6, column=0, columnspan=2, padx=200, pady=5, sticky='we')

def return_to_start():
    global rankings_button
    clear_instructions()
    update_instructions(
        """Please open the Deep Dive Template in Excel and paste the values in the appropriate cells.\n\nInstructions:\n
        1. Paste the vendor rankings to D2 if YTD or J2 if Last Year.
        2. Paste the agent rankings to G2 if YTD or M2 if Last Year.
        3. Save the file and close it.
        4. Click the 'Get New Template' button to start over.
        """
    )

    # Remove the "Finish" button
    rankings_button.grid_forget()

    # Add the "Run" button
    run_button = ttk.Button(root, text="Run", command=on_run)
    run_button.grid(row=6, column=0, columnspan=2, padx=200, pady=5, sticky='we')

def save_excel_template():
    # Get the initial directory
    initial_dir = os.path.join(os.getcwd(), 'public')
    print("Hit 1")

    # Prompt the user to save a file
    filepath = filedialog.asksaveasfilename(initialdir='/Users/jakobbellamy/Documents/Reports',
                                            initialfile='Deep Dive Template.xlsx',
                                            defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
    
    print("Hit 2")
    if not filepath:
        print("No file selected.")
        return

    # Copy the template file to the chosen location
    print("Hit 3")
    template_path = os.path.join(initial_dir, 'Deep Dive Template.xlsx')
    print("Hit 4")
    shutil.copyfile(template_path, filepath)
    print("Hit 5")

    # Show a popup box with more information
messagebox.showinfo("Important!", 
                    "The template file has been saved successfully.\n\n"
                    "Please open the file and use the following instructions:\n"
                    "1. Navigate to the Scratch tab.\n"
                    "2. In B2, type the name of the account as you would like it to appear in the report.\n"
                    "3. In B3, type the month date (ex. January) of the most recent reporting data (should be the same as dropdown).\n"
                    "4. In B4, type the year of the most recent reporting data.\n\n"
                    "It is assumed that data goes through December of the previous year.")

# Set up the main application window
root = tk.Tk()
root.title("Core Explore Kit")

# Set App Window
root.geometry("600x330")

# Add a window title
title_font = font.Font(family="Helvetica", size=20, weight="bold")

title = tk.Label(root, text="Deep Dive Copilot", 
         font=title_font)
title.grid(row=0, column=0, padx=10, pady=5, sticky='we')

# Add a button that prompts the user to save a file
save_button = ttk.Button(root, text="Get New Template", command=save_excel_template)
save_button.grid(row=0, column=1, padx=10, pady=5, sticky='ns')

# Add a button to refresh the file list by processing new core reports 

root.grid_rowconfigure(0, weight=1)


# Add a separator
separator = ttk.Separator(root, orient='horizontal')
separator.grid(row=2, column=0, columnspan=2, sticky='we', pady=10)

# Sequence Selection
tk.Label(root, text="Select Date:").grid(row=3, column=0, padx=10, pady=5, sticky='w')
sequence_var = tk.StringVar()
sequence_dropdown = ttk.Combobox(root, textvariable=sequence_var)

dates_sequence = sorted(list(get_sequences()), reverse=True)

# Optional pairing Down
dates_sequence = [date for date in dates_sequence if date in ['122023', '062024']]

sequence_dropdown['values'] = dates_sequence
sequence_dropdown.grid(row=3, column=1, padx=10, pady=5, sticky='we')
sequence_dropdown.bind('<<ComboboxSelected>>', on_sequence_select)

# File Selection
tk.Label(root, text="Select File:").grid(row=4, column=0, padx=10, pady=5, sticky='w')
file_mapping = get_files()
file_var = tk.StringVar()
file_dropdown = ttk.Combobox(root, textvariable=file_var)
file_dropdown['values'] = sorted(list(get_files()))
file_dropdown.grid(row=4, column=1, padx=10, pady=5, sticky='we')

# Instructions Text
instructions = tk.Text(root, height=10, width=50, state='disabled', bg='white')
instructions.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='we')

update_instructions(
"""Please open your Deep Dive Template in Excel.\n\nInstructions:\n
1. Select a date from the dropdown.
2. Select a file from the dropdown.
3. Click the Run button then follow the instructions in the console.
""")

# Run Button
run_button = ttk.Button(root, text="Run", command=on_run)
run_button.grid(row=6, column=0, columnspan=2, padx=200, pady=5, sticky='we')


# Configure column 1 to expand with window size
root.columnconfigure(1, weight=1)

# Start the Tkinter event loop
root.mainloop()
