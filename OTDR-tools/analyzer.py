import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Scrollbar, Listbox
import subprocess
import os
import json
import pandas as pd
import re
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for PyInstaller bundles. """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Use the function to get the correct path
rbOTDR_path = resource_path("rbOTDR.rb")

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


def apply_conditional_formatting(worksheet, df, workbook):
    """
    Applies conditional formatting to the Comments column or Event_X_Comments columns based on the report type.
    """
    # Formats
    format_critical = workbook.add_format({'bg_color': '#FF0000', 'align': 'left', 'bold': True})  # Red
    format_warning = workbook.add_format({'bg_color': '#FFA500', 'align': 'left', 'bold': True})  # Orange
    format_pass = workbook.add_format({'bg_color': '#D9EAD3', 'align': 'left', 'bold': True})  # Green

    # Check if it is a wide or stacked report
    if "Comments" in df.columns:  # Stacked report
        comments_col_idx = df.columns.get_loc("Comments")  # Get index of the Comments column
        worksheet.conditional_format(
            1, comments_col_idx, len(df), comments_col_idx,  # First data row to last data row
            {'type': 'text', 'criteria': 'containing', 'value': 'Possible break', 'format': format_critical}
        )
        worksheet.conditional_format(
            1, comments_col_idx, len(df), comments_col_idx,  # First data row to last data row
            {'type': 'text', 'criteria': 'containing', 'value': 'Possible microbend', 'format': format_warning}
        )
        worksheet.conditional_format(
            1, comments_col_idx, len(df), comments_col_idx,  # First data row to last data row
            {'type': 'text', 'criteria': 'containing', 'value': 'Pass', 'format': format_pass}
        )
    else:  # Wide report
        for col_idx, col_name in enumerate(df.columns):
            if col_name.endswith("_Comments"):  # Apply formatting to Event_X_Comments columns
                worksheet.conditional_format(
                    1, col_idx, len(df), col_idx,  # First data row to last data row
                    {'type': 'text', 'criteria': 'containing', 'value': 'Possible break', 'format': format_critical}
                )
                worksheet.conditional_format(
                    1, col_idx, len(df), col_idx,  # First data row to last data row
                    {'type': 'text', 'criteria': 'containing', 'value': 'Possible microbend', 'format': format_warning}
                )
                worksheet.conditional_format(
                    1, col_idx, len(df), col_idx,  # First data row to last data row
                    {'type': 'text', 'criteria': 'containing', 'value': 'Pass', 'format': format_pass}
                )

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

        # Retrieve user-defined tolerances
        user_pass_tolerance = pass_tolerance.get()
        user_warning_tolerance = warning_tolerance.get()

        # Analyze the entire run per shot direction and update comments
        for record in extracted_data:
            total_splice_loss = 0
            total_distance = 0
            critical_flag = False
            warning_flag = False

            # Analyze each event in the record
            event_counter = 1
            while f"Event_{event_counter}" in record:
                splice_loss = float(record.get(f"Event_{event_counter}_Splice_Loss", 0) or 0)
                if splice_loss > user_warning_tolerance:  # Critical
                    record[f"Event_{event_counter}_Comments"] = "Possible break"
                    critical_flag = True
                elif splice_loss > user_pass_tolerance:  # Warning
                    record[f"Event_{event_counter}_Comments"] = "Possible microbend"
                    warning_flag = True
                else:  # Pass
                    record[f"Event_{event_counter}_Comments"] = "Pass"
                event_counter += 1

            # Consolidate the overall comment for the shot direction
            if critical_flag:
                record["Comments"] = "Critical issues detected (Possible break)"
            elif warning_flag:
                record["Comments"] = "Warnings detected (Possible microbend)"
            else:
                record["Comments"] = "Pass"

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

                # Adjust the width of each column
                for idx, col in enumerate(df.columns):
                    max_length = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
                    worksheet.set_column(idx, idx, max_length * 1.25, workbook.add_format({'align': 'left'}))

                # Apply conditional formatting
                apply_conditional_formatting(worksheet, df, workbook)

            messagebox.showinfo("Success", f"Report saved to: {save_path}")
        else:
            messagebox.showerror("Cancelled", "Report save cancelled.")

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def create_tolerances_tab(notebook):
    """
    Creates the Tolerances tab where users can input their own tolerances.
    """
    tolerances_tab = ttk.Frame(notebook)
    notebook.add(tolerances_tab, text="Tolerances")

    tk.Label(tolerances_tab, text="Set Tolerances for Fiber Analysis").pack(pady=10)

    # Tolerance input variables
    global pass_tolerance, warning_tolerance
    pass_tolerance = tk.DoubleVar(value=0.3)
    warning_tolerance = tk.DoubleVar(value=0.6)

    # Pass Tolerance
    tk.Label(tolerances_tab, text="Maximum Splice Loss for Pass (dB):").pack(pady=5)
    tk.Entry(tolerances_tab, textvariable=pass_tolerance).pack(pady=5)

    # Warning Tolerance
    tk.Label(tolerances_tab, text="Maximum Splice Loss for Warning (dB):").pack(pady=5)
    tk.Entry(tolerances_tab, textvariable=warning_tolerance).pack(pady=5)

    tk.Label(tolerances_tab, text="Critical issues detected if splice loss exceeds Warning tolerance.").pack(pady=10)

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

def create_json_processing_tab(notebook):
    """
    Creates the JSON Processing tab for the application.
    """
    json_tab = ttk.Frame(notebook)
    notebook.add(json_tab, text="JSON Processing")

    tk.Label(json_tab, text="JSON Processing Tab").pack(pady=5)

    btn_select_files = tk.Button(json_tab, text="Select JSON Files", command=select_files)
    btn_select_files.pack(pady=5)

    btn_generate_report = tk.Button(json_tab, text="Generate Report", command=generate_report)
    btn_generate_report.pack(pady=5)

    json_file_listbox_frame = tk.Frame(json_tab)
    json_file_listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)

    global json_file_listbox
    json_file_listbox = tk.Listbox(json_file_listbox_frame, height=15, width=80)
    json_scrollbar = Scrollbar(json_file_listbox_frame, orient=tk.VERTICAL, command=json_file_listbox.yview)
    json_file_listbox.config(yscrollcommand=json_scrollbar.set)

    json_file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    json_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    global wide_report_var
    wide_report_var = tk.BooleanVar()
    chk_wide_report = tk.Checkbutton(json_tab, text="Generate Wide Report", variable=wide_report_var)
    chk_wide_report.pack(pady=10)

def create_sor_parsing_tab(notebook):
    """
    Creates the .sor Parsing tab for the application.
    """
    sor_tab = ttk.Frame(notebook)
    notebook.add(sor_tab, text=".sor Parsing")

    global sor_input_var, sor_output_var
    sor_input_var = tk.StringVar()
    sor_output_var = tk.StringVar()

    tk.Label(sor_tab, text="Input Folder (Containing .sor Files):").pack(pady=5)
    tk.Entry(sor_tab, textvariable=sor_input_var, width=60).pack(pady=5)
    tk.Button(sor_tab, text="Select Input Folder", command=lambda: sor_input_var.set(filedialog.askdirectory())).pack(pady=5)

    tk.Label(sor_tab, text="Output Folder (For Parsed Files):").pack(pady=5)
    tk.Entry(sor_tab, textvariable=sor_output_var, width=60).pack(pady=5)
    tk.Button(sor_tab, text="Select Output Folder", command=lambda: sor_output_var.set(filedialog.askdirectory())).pack(pady=5)

    def parse_sor_files():
        """
        Parses .sor files using the Ruby script.
        """
        input_folder = sor_input_var.get()
        output_folder = sor_output_var.get()
        rbOTDR_path = resource_path("rbOTDR.rb")

        if not os.path.isdir(input_folder) or not os.path.isdir(output_folder):
            messagebox.showerror("Error", "Select valid input and output folders.")
            return

        error_log = []

        for root, _, files in os.walk(input_folder):
            for file in files:
                if file.endswith(".sor"):
                    sor_file_path = os.path.join(root, file)
                    try:
                        result = subprocess.run(
                            ["ruby", rbOTDR_path, sor_file_path, output_folder],
                            check=True,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            text=True
                        )
                        print(f"Successfully parsed: {sor_file_path}")
                    except FileNotFoundError:
                        messagebox.showerror("Error", "Ruby interpreter not found. Ensure Ruby is installed and added to PATH.")
                        return
                    except subprocess.CalledProcessError as e:
                        error_message = (
                            f"Failed to parse {sor_file_path}.\n"
                            f"Exit Code: {e.returncode}\n"
                            f"Output: {e.stdout}\n"
                            f"Error: {e.stderr}"
                        )
                        print(error_message)
                        error_log.append(error_message)
                    except Exception as e:
                        error_message = f"Unexpected error with {sor_file_path}: {e}"
                        print(error_message)
                        error_log.append(error_message)

        if error_log:
            error_log_path = os.path.join(output_folder, "error_log.txt")
            with open(error_log_path, "w") as log_file:
                log_file.write("\n".join(error_log))
            messagebox.showerror("Error", f"Some files failed to parse. See log at: {error_log_path}")
        else:
            messagebox.showinfo("Success", ".sor files parsed successfully.")

    # Add Parse Button
    btn_parse_sor = tk.Button(sor_tab, text="Parse .sor Files", command=parse_sor_files)
    btn_parse_sor.pack(pady=10)

def main():
    # Main application window
    window = tk.Tk()
    window.title("JSON and .sor File Analyzer")
    window.geometry("800x600")

    # Create notebook for tabs
    notebook = ttk.Notebook(window)
    notebook.pack(fill=tk.BOTH, expand=True)

    # Add JSON Processing, .sor Parsing, and Tolerances tabs
    create_json_processing_tab(notebook)
    create_sor_parsing_tab(notebook)
    create_tolerances_tab(notebook)

    # Run the application
    window.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()
    finally:
        input("Press Enter to exit...")

