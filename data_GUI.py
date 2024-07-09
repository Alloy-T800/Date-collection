import re
import pandas as pd
import tkinter as tk
import numpy as np
import tkinter.font as tkFont
from tkinter import messagebox, scrolledtext
import shutil
import os
import glob
from datetime import datetime

data_path = "data_set.xlsx"

def backup_excel_file(original_file, max_backups=10):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = original_file.replace('.xlsx', f'_backup_{timestamp}.xlsx')
    shutil.copyfile(original_file, backup_file)

    # Get all backup files and sort by modification time
    backup_files = sorted(glob.glob(original_file.replace('.xlsx', '_backup_*.xlsx')), key=os.path.getmtime)

    # Keep the latest max_backups backups, delete the rest
    for old_backup in backup_files[:-max_backups]:
        os.remove(old_backup)

# Alloy automatic recognition and processing
def parse_alloy_composition(alloy_str):
    # Remove spaces and initialize dictionary
    alloy_str = alloy_str.replace(" ", "")
    element_counts = {}

    # Regular expression to separate elements inside and outside parentheses
    pattern = r'\((?P<inner>[A-Za-z\d\.]+)\)(?P<inner_multiplier>\d*)|(?P<element>[A-Z][a-z]*)(?P<count>\d*\.?\d*)'
    matches = re.findall(pattern, alloy_str)

    # Total amounts calculation inside and outside parentheses
    inner_multiplier = 1  # Default to 1 for cases with no number outside parentheses
    outer_total = 0
    inner_composition = ""

    for inner, inner_mult, element, count in matches:
        if inner:
            inner_composition = inner
            inner_multiplier = float(inner_mult) if inner_mult else 1  # Handle cases with no number outside parentheses
        else:
            count = float(count) if count else 1
            outer_total += count
            element_counts[element] = count

    # Calculate proportions of elements inside parentheses
    inner_matches = re.findall(r'([A-Z][a-z]*)(\d*\.?\d*)', inner_composition)
    inner_total = sum([float(c) if c else 1 for _, c in inner_matches])

    for e, c in inner_matches:
        c = float(c) if c else 1
        element_proportion = c / inner_total
        element_counts[e] = element_counts.get(e, 0) + element_proportion * inner_multiplier

    # Adjust proportions of all elements
    total_amount = sum(element_counts.values())
    for element in element_counts:
        element_counts[element] = round((element_counts[element] / total_amount) * 100, 4)

    return element_counts

# Performance input and validation
def process_property_input(property_value):
    if property_value.strip() == "":
        return None
    try:
        return float(property_value)
    except ValueError:
        return "Error: Please enter a valid number."

# Phase properties input
def process_phase_properties(phase_selections):
    phases = ["FCC", "BCC", "L12", "B2", "L21", "Laves", "σ", "Im"]
    phase_properties = {}

    # Check if there is any "Yes" selection
    any_yes = any(choice == 'Yes' for choice in phase_selections.values())

    for phase in phases:
        if any_yes:
            # For new alloys, if there is any "Yes" selection, set to 1 or 0
            phase_properties[phase] = 1 if phase_selections.get(phase) == 'Yes' else 0
        else:
            # If no "Yes" selection, keep as None
            phase_properties[phase] = None

    return phase_properties

# Custom properties input
def process_custom_properties(eutectic_input, doi_input):
    return {
        "Eutectic": eutectic_input if eutectic_input.strip() != "" else None,
        "Doi": doi_input if doi_input.strip() != "" else None
    }

# Check if the alloy already exists
def check_if_alloy_exists(df, alloy_composition_parsed):

    alloy_exists = df[df[list(alloy_composition_parsed.keys())].eq(pd.Series(alloy_composition_parsed)).all(axis=1)]
    return not alloy_exists.empty, alloy_exists

# Read Excel file
def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

# Update data count label
def update_data_count_label(df, label):
    if df is not None:
        count = len(df)
        label.config(text=f"Data Count: {count}")
    else:
        label.config(text="Data Count: 0")

# Write to Excel file
def write_excel(df, file_path):
    # Backup first
    backup_excel_file(file_path)

    # Then write data
    try:
        df.to_excel(file_path, index=False)
        print("File saved successfully")
    except Exception as e:
        print(f"Error writing file: {e}")

    # Update data count display
    update_data_count_label(df, data_count_label)

# Open a new window for user to input performance data
def open_new_window_for_input(alloy_data, alloy_composition_parsed, result_area, alloy_entry, new_entry=True):
    input_window = tk.Toplevel()
    input_window.title("Input Performance Data")

    # Define fonts
    chinese_font = tkFont.Font(family="楷体", size=14, weight="bold")
    english_font = tkFont.Font(family="Times New Roman", size=14, weight="bold")

    # Create input fields for performance properties
    properties = ["T_YS", "T_US", "T_Elo", "C_YS", "C_Elo", "Hardness", "Ecorr", "Epit"]
    entries = {}
    for idx, prop in enumerate(properties):
        tk.Label(input_window, text=prop + ":", font=chinese_font).grid(row=idx, column=0, sticky='w', padx=10, pady=5)
        entry = tk.Entry(input_window, font=english_font)
        entry.grid(row=idx, column=1, padx=10, pady=5)
        entries[prop] = entry

    # Add buttons for FCC, BCC, etc. properties
    phase_vars = {}
    phases = ["FCC", "BCC", "L12", "B2", "L21", "Laves", "σ", "Im"]
    for idx, phase in enumerate(phases, start=len(properties)):
        tk.Label(input_window, text=phase + ":", font=chinese_font).grid(row=idx, column=0, sticky='w', padx=10, pady=5)
        var = tk.StringVar(value='Not Selected' if new_entry else None)
        yes_button = tk.Radiobutton(input_window, text="Yes", variable=var, value='Yes', font=chinese_font)
        no_button = tk.Radiobutton(input_window, text="No", variable=var, value='No', font=chinese_font)
        none_button = tk.Radiobutton(input_window, text="Not Selected", variable=var, value='Not Selected', font=chinese_font)
        yes_button.grid(row=idx, column=1, padx=10, pady=5)
        no_button.grid(row=idx, column=2, padx=10, pady=5)
        none_button.grid(row=idx, column=3, padx=10, pady=5)
        phase_vars[phase] = var

    # Add input fields for Eutectic and Doi
    eutectic_entry = tk.Entry(input_window, font=english_font)
    doi_entry = tk.Entry(input_window, font=english_font)
    eutectic_entry.grid(row=len(phases) + len(properties), column=1, padx=10, pady=5)
    doi_entry.grid(row=len(phases) + len(properties) + 1, column=1, padx=10, pady=5)
    tk.Label(input_window, text="Eutectic:", font=chinese_font).grid(row=len(phases) + len(properties), column=0, sticky='w', padx=10, pady=5)
    tk.Label(input_window, text="Doi:", font=chinese_font).grid(row=len(phases) + len(properties) + 1, column=0, sticky='w', padx=10, pady=5)

    # Submit button
    submit_button = tk.Button(input_window, text="Submit", font=chinese_font, command=lambda: submit_performance_data(
        entries, alloy_data, alloy_composition_parsed, result_area, input_window, alloy_entry, phase_vars,
        eutectic_entry, doi_entry))
    submit_button.grid(row=len(phases) + len(properties) + 2, column=1, padx=10, pady=5)

# Update alloy data window
def save_updated_alloy_data(entries, phase_vars, eutectic_entry, doi_entry, existing_alloys, result_area, update_window, alloy_entry):
    global df
    alloy_index = existing_alloys.index[0]  # Assume only dealing with the first matching alloy

    # Extract and update performance data
    for prop, entry in entries.items():
        new_value = process_property_input(entry.get())
                if new_value != existing_alloys.at[alloy_index, prop]:
            df.loc[alloy_index, prop] = new_value

    # Update phase properties
    for phase, var in phase_vars.items():
        df.loc[alloy_index, phase] = 1 if var.get() == '1' else 0

    # Update Eutectic and Doi
    df.loc[alloy_index, "Eutectic"] = eutectic_entry.get() if eutectic_entry.get() else df.loc[alloy_index, "Eutectic"]
    df.loc[alloy_index, "Doi"] = doi_entry.get() if doi_entry.get() else df.loc[alloy_index, "Doi"]

    # Save updates to Excel
    write_excel(df, data_path)

    # Reload dataset
    df = read_excel(data_path)

    # Update result area in main window
    updated_info = df.loc[alloy_index].to_string()
    result_area.delete('1.0', tk.END)
    result_area.insert(tk.INSERT, "Updated alloy data:\n" + updated_info)

    # Clear alloy composition entry
    alloy_entry.delete(0, tk.END)

    # Close update window
    update_window.destroy()

# Window for handling existing alloys
def open_window_for_existing_alloy(alloy_data, existing_alloys, result_area, alloy_entry):
    update_window = tk.Toplevel()
    update_window.title("Update Alloy Information")

    # Display existing alloy information
    existing_info = existing_alloys.iloc[0]  # Assume only dealing with one duplicate alloy
    row_idx = 0
    # Define fonts
    chinese_font = tkFont.Font(family="楷体", size=14, weight="bold")
    english_font = tkFont.Font(family="Times New Roman", size=14, weight="bold")

    # Create input fields for performance properties and initialize with existing values
    entries = {}
    for prop in ["T_YS", "T_US", "T_Elo", "C_YS", "C_Elo", "Hardness", "Ecorr", "Epit"]:
        tk.Label(update_window, text=prop + ":", font=chinese_font).grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
        entry = tk.Entry(update_window, font=english_font)
        entry.insert(0, existing_info[prop])  # Initialize with existing values
        entry.grid(row=row_idx, column=1)
        entries[prop] = entry
        row_idx += 1

    # Create radio buttons for phase properties and initialize with existing values
    phase_vars = {}
    for phase in ["FCC", "BCC", "L12", "B2", "L21", "Laves", "σ", "Im"]:
        tk.Label(update_window, text=phase + ":", font=chinese_font).grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
        # Check for NaN values and handle accordingly
        if np.isnan(existing_info[phase]):
            phase_value = 'Not Selected'
        else:
            phase_value = str(int(existing_info[phase]))

        var = tk.StringVar(value=phase_value)
        yes_button = tk.Radiobutton(update_window, text="Yes", variable=var, value='1', font=chinese_font)
        no_button = tk.Radiobutton(update_window, text="No", variable=var, value='0', font=chinese_font)

        yes_button.grid(row=row_idx, column=1)
        no_button.grid(row=row_idx, column=2)
        phase_vars[phase] = var
        row_idx += 1

    # Create input fields for Eutectic and Doi and initialize with existing values
    eutectic_entry = tk.Entry(update_window, font=english_font)
    doi_entry = tk.Entry(update_window, font=english_font)
    eutectic_entry.insert(0, existing_info["Eutectic"])
    doi_entry.insert(0, existing_info["Doi"])
    eutectic_entry.grid(row=row_idx, column=1, padx=10, pady=5)
    doi_entry.grid(row=row_idx + 1, column=1, padx=10, pady=5)
    tk.Label(update_window, text="Eutectic:", font=chinese_font).grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
    tk.Label(update_window, text="Doi:", font=chinese_font).grid(row=row_idx + 1, column=0, sticky='w', padx=10, pady=5)

    # Save button
    save_button = tk.Button(update_window, text="Save", command=lambda: save_updated_alloy_data(
        entries, phase_vars, eutectic_entry, doi_entry, existing_alloys, result_area, update_window, alloy_entry))
    save_button.grid(row=row_idx + 2, column=1)

def submit_performance_data(entries, alloy_data, alloy_composition_parsed, result_area, input_window, alloy_entry, phase_vars, eutectic_entry, doi_entry):
    global df
    performance_data = {}
    valid_input = True

    for prop, entry in entries.items():
        property_value = process_property_input(entry.get())
        if isinstance(property_value, str):  # Check for error messages
            messagebox.showerror("Input Error", f"{prop}: {property_value}")
            valid_input = False
            break
        performance_data[prop] = property_value

    if valid_input:
        # Ensure all unspecified elements are set to 0
        all_elements = set(['Fe', 'Ni', 'Cr', 'Al', 'Ti', 'Mo', 'Nb', 'V', 'Co', 'Mn', 'Cu', 'Zr', 'Hf', 'Ta', 'Si', 'W'])
        for element in all_elements - set(alloy_composition_parsed.keys()):
            alloy_composition_parsed[element] = 0

        # Process phase properties input
        phase_selections = {phase: var.get() for phase, var in phase_vars.items()}
        phase_properties = process_phase_properties(phase_selections)

        # Process Eutectic and Doi input
        custom_properties = process_custom_properties(eutectic_entry.get(), doi_entry.get())

        # Integrate all data
        alloy_data_row = {"Alloy": alloy_data}
        alloy_data_row.update(alloy_composition_parsed)
        alloy_data_row.update(performance_data)
        alloy_data_row.update(phase_properties)
        alloy_data_row.update(custom_properties)

        # Write data to Excel sheet
        df = read_excel(data_path)  # Read existing data
        new_df = pd.DataFrame([alloy_data_row])  # Create new data row
        updated_df = pd.concat([df, new_df], ignore_index=True)
        write_excel(updated_df, data_path)  # Write updated data

        # Reload dataset
        df = read_excel(data_path)

        # Clear alloy composition entry
        alloy_entry.delete(0, tk.END)

        # Close input window
        input_window.destroy()

        # Update result area in main window
        result_area.delete('1.0', tk.END)
        result_str = "\n".join(f"{k}: {v}" for k, v in alloy_data_row.items())
        result_area.insert(tk.INSERT, "Saved alloy data:\n" + result_str)

        # Close input window
        input_window.destroy()

def main():
    global data_count_label  # Declare as global variable

    # Create main window
    root = tk.Tk()
    root.title("Alloy Composition and Performance Input")
    root.geometry("800x850")

    # Define fonts
    chinese_font = tkFont.Font(family="楷体", size=14, weight="bold")
    english_font = tkFont.Font(family="Times New Roman", size=14, weight="bold")

    # Alloy composition input label and entry
    label_alloy = tk.Label(root, text="Alloy Composition:", font=chinese_font)
    label_alloy.grid(row=0, column=0, padx=10, pady=10)
    alloy_entry = tk.Entry(root, font=english_font, width=50)
    alloy_entry.grid(row=0, column=1, padx=10, pady=10)

    # Result display area
    result_area = scrolledtext.ScrolledText(root, height=30, width=70, font=english_font)
    result_area.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    # Submit button
    submit_button = tk.Button(root, text="Submit", font=chinese_font,
                              command=lambda: submit_alloy_data(alloy_entry.get(), result_area, alloy_entry))
    submit_button.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

    # Data count label
    data_count_label = tk.Label(root, text="Data Count: 0", font=chinese_font)
    data_count_label.grid(row=3, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

    # Load dataset and update label
    global df
    df = read_excel(data_path)
    update_data_count_label(df, data_count_label)

    root.mainloop()

def dataframe_to_formatted_string(df):
    formatted_str = ""
    for index, row in df.iterrows():
        row_str = f"Row {index}:\n"
        for col in df.columns:
            row_str += f"    {col}: {row[col]}\n"
        formatted_str += row_str + "\n"
    return formatted_str

def submit_alloy_data(alloy_data, result_area, alloy_entry):
    alloy_composition_parsed = parse_alloy_composition(alloy_data)
    exists, existing_alloys = check_if_alloy_exists(df, alloy_composition_parsed)

    if exists:
        existing_info = dataframe_to_formatted_string(existing_alloys)
        result_area.delete('1.0', tk.END)
        result_area.insert(tk.INSERT, f"Duplicate alloy information:\n{existing_info}")
        # Open update window
        open_window_for_existing_alloy(alloy_data, existing_alloys, result_area, alloy_entry)
    else:
        open_new_window_for_input(alloy_data, alloy_composition_parsed, result_area, alloy_entry)

if __name__ == "__main__":
    main()

