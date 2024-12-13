import tkinter as tk
from tkinter import filedialog, messagebox, Scrollbar, Listbox
import json
import pandas as pd
import re


def process_json_and_extract(file_paths):
    """
    Processes JSON files and extracts structured data.
    Returns a list of dictionaries with extracted fields to match the format of the XLSX report.
    """
    extracted_data = []
    for file_path in file_paths:
        try:
            with open(file_path, 'r') as file:
                json_data = json.load(file)

                # Extract the shot direction and fiber ID from filename
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
                            "Range": range_value,  # Add the range column next to event distance
                            "Splice_Loss": event_info.get("splice loss", ""),
                            "Refl_Loss": event_info.get("refl loss", ""),
                            "Comments": event_info.get("comments", ""),
                        })
                
                # Total shot distance (highest distance found in the events)
                distance_km = max_distance if not distance_km else distance_km

                # Prepare the record to match the format for the report
                record = {
                    "Shot_Direction": shot_direction,  # Added shot direction
                    "Fiber_ID": fiber_id,
                    "Distance_KM": distance_km,  # Total shot distance
                }

                # Add event data as rows under the fiber ID block
                for event in event_data:
                    event_record = {
                        "Shot_Direction": shot_direction,  # Added shot direction
                        "Fiber_ID": fiber_id,
                        "Event": event["Event"],
                        "Distance": event["Distance"] if event["Distance"] != 0 else "",  # Avoid 0 values
                        "Range": event["Range"],  # Add the range value here
                        "Splice_Loss": event["Splice_Loss"] if event["Splice_Loss"] != 0 else "",  # Avoid 0 values
                        "Refl_Loss": event["Refl_Loss"] if event["Refl_Loss"] != 0 else "",  # Avoid 0 values
                        "Comments": event["Comments"] if event["Comments"] != "" else "",  # Avoid empty text
                    }
                    extracted_data.append(event_record)

                # Add an empty row after each fiber's data block for readability
                extracted_data.append({
                    "Shot_Direction": "",  # Empty cell for separation
                    "Fiber_ID": "",
                    "Event": "",
                    "Distance": "",
                    "Range": "",
                    "Splice_Loss": "",
                    "Refl_Loss": "",
                    "Comments": "",
                })

        except Exception as e:
            print(f"Error processing {file_path}: {e}")
    return extracted_data


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
    Processes the selected JSON files, generates a report, and saves it as an Excel file.
    """
    if not window.file_paths:
        messagebox.showerror("Error", "No files selected for processing.")
        return

    try:
        extracted_data = process_json_and_extract(window.file_paths)
        
        if not extracted_data:
            messagebox.showerror("Error", "No valid data extracted.")
            return

        # Create DataFrame from extracted JSON data
        df = pd.DataFrame(extracted_data)

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

                # Create a left-aligned format for all columns
                left_aligned_format = workbook.add_format({'align': 'left'})

                # Adjust the width of each column by adding 50% more space
                for idx, col in enumerate(df.columns):
                    max_length = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
                    worksheet.set_column(idx, idx, max_length * 1.5, left_aligned_format)  # 50% extra width

            messagebox.showinfo("Success", f"Report saved to: {save_path}")
        else:
            messagebox.showerror("Cancelled", "Report save cancelled.")

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")


# Set up the main application window
window = tk.Tk()
window.title("JSON File Analyzer")
window.geometry("700x500")

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
