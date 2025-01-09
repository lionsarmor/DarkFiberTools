import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Scrollbar, Listbox
import subprocess
import os
import json
import pandas as pd
import re

def process_json_and_extract(file_paths):
    """
    Processes JSON files and extracts structured data for both stacked and wide report formats.
    Returns a list of dictionaries with extracted fields to match the format of the XLSX report.
    """
    extracted_data = []

    for file_path in file_paths:
        try:
            with open(file_path, 'r') as file:
                json_data = json.load(file)

                # Extract shot direction and fiber ID from filename
                filename = json_data.get("filename", "")
                shot_direction = " ".join(filename.split(" ")[:2])  # First two parts of the filename
                fiber_id = re.search(r"\s(\d{3})\s", filename)
                fiber_id = fiber_id.group(1) if fiber_id else ""

                # Extract fields from JSON
                distance_km = json_data.get("GenParams", {}).get("distance_km", "")
                range_value = json_data.get("FxdParams", {}).get("range", "")  # Extract range value

                # Get all events
                events = json_data.get("KeyEvents", {})
                event_data = []
                max_distance = 0

                for event_key, event_info in events.items():
                    if event_key.startswith("event"):
                        distance = float(event_info.get("distance", 0))
                        if distance > max_distance:
                            max_distance = distance
                        event_data.append({
                            "Event": event_key,
                            "Distance": distance,
                            "Splice_Loss": event_info.get("splice loss", ""),
                            "Refl_Loss": event_info.get("refl loss", ""),
                            "Comments": event_info.get("comments", ""),
                        })

                # Total shot distance (highest distance found in the events)
                distance_km = max_distance if not distance_km else distance_km

                # Prepare the record to match both wide and stacked formats
                record = {
                    "Shot_Direction": shot_direction,
                    "Fiber_ID": fiber_id,
                    "Range": range_value,  # Range will be on the left side, alongside Fiber_ID
                    "Distance_KM": distance_km,  # Total shot distance
                }

                # Add event data as columns under the fiber ID block (for wide format)
                event_counter = 1
                for event in event_data:
                    record[f"Event_{event_counter}"] = event["Event"]
                    record[f"Event_{event_counter}_Distance"] = event["Distance"]
                    record[f"Event_{event_counter}_Splice_Loss"] = event["Splice_Loss"]
                    record[f"Event_{event_counter}_Refl_Loss"] = event["Refl_Loss"]
                    record[f"Event_{event_counter}_Comments"] = event["Comments"]
                    event_counter += 1

                extracted_data.append(record)

        except Exception as e:
            print(f"Error processing {file_path}: {e}")

    return extracted_data

def generate_stacked_report(extracted_data):
    """
    Generates the stacked report where events are stacked in rows.
    This will add a blank row between different Fiber IDs to separate them.
    """
    stacked_data = []
    last_fiber_id = None

    for record in extracted_data:
        shot_direction = record['Shot_Direction']
        fiber_id = record['Fiber_ID']
        range_value = record['Range']
        distance_km = record['Distance_KM']

        # If fiber ID changes, add a blank row to separate fibers
        if fiber_id != last_fiber_id:
            if last_fiber_id is not None:  # Skip the blank row before the first fiber
                stacked_data.append({
                    "Shot_Direction": "",
                    "Fiber_ID": "",
                    "Range": "",
                    "Distance_KM": "",
                    "Event": "",
                    "Event_Distance": "",
                    "Splice_Loss": "",
                    "Refl_Loss": "",
                    "Comments": ""
                })
            last_fiber_id = fiber_id

        # For each event, create a new row
        event_counter = 1
        while f"Event_{event_counter}" in record:
            event_data = {
                "Shot_Direction": shot_direction,
                "Fiber_ID": fiber_id,
                "Range": range_value,
                "Distance_KM": distance_km,
                "Event": record[f"Event_{event_counter}"],
                "Event_Distance": record[f"Event_{event_counter}_Distance"],
                "Splice_Loss": record[f"Event_{event_counter}_Splice_Loss"],
                "Refl_Loss": record[f"Event_{event_counter}_Refl_Loss"],
                "Comments": record[f"Event_{event_counter}_Comments"]
            }
            stacked_data.append(event_data)
            event_counter += 1

    return pd.DataFrame(stacked_data)

def generate_wide_report(extracted_data):
    """
    Generates the wide report where events are placed on the X-axis.
    """
    # Create DataFrame from extracted JSON data
    df = pd.DataFrame(extracted_data)

    # In the wide format, each event should have its own columns (Event, Distance, Splice Loss, etc.)
    wide_report_data = []
    for idx, row in df.iterrows():
        shot_direction = row['Shot_Direction']
        fiber_id = row['Fiber_ID']
        range_value = row['Range']
        distance_km = row['Distance_KM']

        # Create a record for this fiber with all events in columns
        event_data = {
            "Shot_Direction": shot_direction,
            "Fiber_ID": fiber_id,
            "Range": range_value,  # Range appears only once on the left side
            "Distance_KM": distance_km,
        }

        event_counter = 1
        while f"Event_{event_counter}" in row:
            event_data[f"Event_{event_counter}"] = row[f"Event_{event_counter}"]
            event_data[f"Event_{event_counter}_Distance"] = row[f"Event_{event_counter}_Distance"]
            event_data[f"Event_{event_counter}_Splice_Loss"] = row[f"Event_{event_counter}_Splice_Loss"]
            event_data[f"Event_{event_counter}_Refl_Loss"] = row[f"Event_{event_counter}_Refl_Loss"]
            event_data[f"Event_{event_counter}_Comments"] = row[f"Event_{event_counter}_Comments"]
            event_counter += 1

        wide_report_data.append(event_data)

    return pd.DataFrame(wide_report_data)

def select_files():
    """
    Opens a file dialog to select multiple JSON files.
    """
    file_paths = filedialog.askopenfilenames(
        title="Select JSON files",
        filetypes=[("JSON Files", "*.json")],
    )
    if file_paths:
        json_file_listbox.delete(0, tk.END)  # Clear previous entries
        for path in file_paths:
            json_file_listbox.insert(tk.END, path)

def generate_report():
    """
    Processes the selected JSON files, generates either a wide or stacked report, and saves it as an Excel file.
    """
    file_paths = list(json_file_listbox.get(0, tk.END))

    if not file_paths:
        messagebox.showerror("Error", "No files selected for processing.")
        return

    try:
        extracted_data = process_json_and_extract(file_paths)

        if not extracted_data:
            messagebox.showerror("Error", "No valid data extracted.")
            return

        # Create DataFrame from extracted JSON data
        if wide_report_var.get():
            df = generate_wide_report(extracted_data)
        else:
            df = generate_stacked_report(extracted_data)

        # Open file dialog to save the report
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Report"
        )

        if save_path:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')

                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # Adjust the width of each column based on its contents plus 25%
                for idx, col in enumerate(df.columns):
                    max_length = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
                    worksheet.set_column(idx, idx, max_length * 1.25)

            messagebox.showinfo("Success", f"Report saved to: {save_path}")
        else:
            messagebox.showerror("Cancelled", "Report save cancelled.")

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

# Main application window
window = tk.Tk()
window.title("JSON and .sor File Analyzer")
window.geometry("800x600")

# Create notebook for tabs
notebook = ttk.Notebook(window)
notebook.pack(fill=tk.BOTH, expand=True)

# Tab 1: JSON Processing
json_tab = ttk.Frame(notebook)
notebook.add(json_tab, text="JSON Processing")

# Widgets for JSON Processing
tk.Label(json_tab, text="JSON Processing Tab").pack(pady=5)

btn_select_files = tk.Button(json_tab, text="Select JSON Files", command=select_files)
btn_select_files.pack(pady=5)

btn_generate_report = tk.Button(json_tab, text="Generate Report", command=generate_report)
btn_generate_report.pack(pady=5)

json_file_listbox_frame = tk.Frame(json_tab)
json_file_listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)

json_file_listbox = tk.Listbox(json_file_listbox_frame, height=15, width=80)
json_scrollbar = Scrollbar(json_file_listbox_frame, orient=tk.VERTICAL, command=json_file_listbox.yview)
json_file_listbox.config(yscrollcommand=json_scrollbar.set)

json_file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
json_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Checkboxes for report format
wide_report_var = tk.BooleanVar()
chk_wide_report = tk.Checkbutton(json_tab, text="Generate Wide Report", variable=wide_report_var)
chk_wide_report.pack(pady=10)

# Tab 2: .sor Parsing
sor_tab = ttk.Frame(notebook)
notebook.add(sor_tab, text=".sor Parsing")

sor_input_var = tk.StringVar()
sor_output_var = tk.StringVar()

tk.Label(sor_tab, text="Input Folder (Containing .sor Files):").pack(pady=5)
tk.Entry(sor_tab, textvariable=sor_input_var, width=60).pack(pady=5)
tk.Button(sor_tab, text="Select Input Folder", command=lambda: sor_input_var.set(filedialog.askdirectory())).pack(pady=5)

tk.Label(sor_tab, text="Output Folder (For Parsed Files):").pack(pady=5)
tk.Entry(sor_tab, textvariable=sor_output_var, width=60).pack(pady=5)
tk.Button(sor_tab, text="Select Output Folder", command=lambda: sor_output_var.set(filedialog.askdirectory())).pack(pady=5)

def parse_sor_files():
    input_folder = sor_input_var.get()
    output_folder = sor_output_var.get()
    rbOTDR_path = "./rbOTDR.rb"

    if not os.path.isdir(input_folder) or not os.path.isdir(output_folder):
        messagebox.showerror("Error", "Select valid input and output folders.")
        return

    for root, _, files in os.walk(input_folder):
        for file in files:
            if file.endswith(".sor"):
                try:
                    subprocess.run([rbOTDR_path, os.path.join(root, file), output_folder], check=True)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to parse {file}: {e}")
    messagebox.showinfo("Success", ".sor files parsed successfully.")

btn_parse_sor = tk.Button(sor_tab, text="Parse .sor Files", command=parse_sor_files)
btn_parse_sor.pack(pady=10)

# Run the application
window.mainloop()
