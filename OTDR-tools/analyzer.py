import tkinter as tk
from tkinter import filedialog, messagebox, Scrollbar, Listbox
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
        window.file_paths = file_paths
        update_file_list(file_paths)


def update_file_list(file_paths):
    """
    Updates the scrollable list window with the names of uploaded files.
    """
    file_listbox.delete(0, tk.END)  # Clear previous entries
    for path in file_paths:
        file_listbox.insert(tk.END, path)


def generate_report():
    """
    Processes the selected JSON files, generates either a wide or stacked report, and saves it as an Excel file.
    """
    if not window.file_paths:
        messagebox.showerror("Error", "No files selected for processing.")
        return

    try:
        extracted_data = process_json_and_extract(window.file_paths)
        
        if not extracted_data:
            messagebox.showerror("Error", "No valid data extracted.")
            return

        # Create DataFrame from extracted JSON data based on the checkbox value
        if wide_report_var.get():
            df = generate_wide_report(extracted_data)  # Generate wide report
        else:
            df = generate_stacked_report(extracted_data)  # Default stacked report

        # Open file dialog to save the report
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Report"
        )

        if save_path:
            # Save the DataFrame to an Excel file with xlsxwriter engine
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')

                # Access the xlsxwriter workbook and worksheet
                workbook  = writer.book
                worksheet = writer.sheets['Sheet1']

                # Create formats for colors
                highlight_format = workbook.add_format({'bg_color': '#D9EAD3', 'align': 'left'})
                red_format = workbook.add_format({'bg_color': '#FF0000', 'color': '#FFFFFF', 'align': 'left'})
                green_format = workbook.add_format({'bg_color': '#008000', 'color': '#FFFFFF', 'align': 'left'})
                left_aligned_format = workbook.add_format({'align': 'left'})

                # Apply color to specific columns
                for col_num, col in enumerate(df.columns):
                    # Color the Shot Direction, Fiber ID, Range columns
                    if col in ["Shot_Direction", "Fiber_ID", "Range"]:
                        worksheet.set_column(col_num, col_num, None, highlight_format)

                    # Color Splice Loss and Refl Loss columns based on value
                    if "Splice_Loss" in col or "Refl_Loss" in col:
                        worksheet.set_column(col_num, col_num, None, red_format)

                    # Highlight Distance columns
                    if "Distance" in col:
                        worksheet.set_column(col_num, col_num, None, green_format)

                    # Left-align all text in every column
                    worksheet.set_column(col_num, col_num, None, left_aligned_format)

                # Adjust the width of each column by adding 50% more space
                for idx, col in enumerate(df.columns):
                    max_length = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
                    worksheet.set_column(idx, idx, max_length * 1.5)

            messagebox.showinfo("Success", f"Report saved to: {save_path}")
        else:
            messagebox.showerror("Cancelled", "Report save cancelled.")

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")


# Set up the main application window
window = tk.Tk()
window.title("JSON File Analyzer")
window.geometry("700x500")

# Checkboxes for report format
wide_report_var = tk.BooleanVar()
chk_wide_report = tk.Checkbutton(window, text="Generate Wide Report", variable=wide_report_var)
chk_wide_report.pack(pady=10)

# Buttons
btn_select_files = tk.Button(window, text="Select JSON Files", command=select_files)
btn_select_files.pack(pady=5)

btn_generate_report = tk.Button(window, text="Generate Report", command=generate_report)
btn_generate_report.pack(pady=5)

# Scrollable Listbox to show selected JSON file names
file_listbox_frame = tk.Frame(window)
file_listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)

file_listbox = tk.Listbox(file_listbox_frame, height=15, width=80)
scrollbar = Scrollbar(file_listbox_frame, orient=tk.VERTICAL, command=file_listbox.yview)
file_listbox.config(yscrollcommand=scrollbar.set)

file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

window.file_paths = []  # Initialize an empty list to store selected file paths

window.mainloop()
