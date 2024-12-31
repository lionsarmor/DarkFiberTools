import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import os


# Functions for JSON Processing (Placeholders for simplicity)
def select_files():
    file_paths = filedialog.askopenfilenames(
        title="Select JSON files",
        filetypes=[("JSON Files", "*.json")],
    )
    if file_paths:
        json_file_listbox.delete(0, tk.END)  # Clear previous entries
        for path in file_paths:
            json_file_listbox.insert(tk.END, path)


def generate_report():
    messagebox.showinfo("Info", "Report generation functionality here.")


# Functions for .sor Parsing
def select_sor_input_folder():
    folder_path = filedialog.askdirectory(title="Select Folder Containing .sor Files")
    if folder_path:
        sor_input_var.set(folder_path)


def select_sor_output_folder():
    folder_path = filedialog.askdirectory(title="Select Output Folder for Parsed Files")
    if folder_path:
        sor_output_var.set(folder_path)


def select_rbotdr_executable():
    file_path = filedialog.askopenfilename(
        title="Select rbOTDR Executable",
        filetypes=[("Executable Files", "*.*")]
    )
    if file_path:
        rbotdr_path_var.set(file_path)


def parse_sor_files():
    input_folder = sor_input_var.get().strip()
    output_folder = sor_output_var.get().strip()
    rbOTDR_path = rbotdr_path_var.get().strip()

    if not os.path.isdir(input_folder):
        messagebox.showerror("Error", "Invalid input folder. Please select a valid folder containing .sor files.")
        return

    if not os.path.isdir(output_folder):
        os.makedirs(output_folder)
        messagebox.showinfo("Info", f"Created output folder: {output_folder}")

    if not os.path.isfile(rbOTDR_path) or not os.access(rbOTDR_path, os.X_OK):
        messagebox.showerror("Error", "Invalid rbOTDR path. Please select a valid executable.")
        return

    try:
        for root, _, files in os.walk(input_folder):
            for file in files:
                if file.lower().endswith('.sor'):
                    input_file_path = os.path.join(root, file)
                    output_file_path = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.json")

                    try:
                        with open(output_file_path, 'w') as output_file:
                            subprocess.run([rbOTDR_path, input_file_path], stdout=output_file, check=True)
                    except subprocess.CalledProcessError as e:
                        messagebox.showerror("Error", f"Error processing {input_file_path}: {e}")
                    except Exception as e:
                        messagebox.showerror("Error", f"Unexpected error: {e}")

        messagebox.showinfo("Success", "All .sor files have been successfully parsed.")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")


# Set up the main application window
window = tk.Tk()
window.title("JSON and .sor File Analyzer")
window.geometry("800x600")

# Create notebook for tabs
notebook = ttk.Notebook(window)
notebook.pack(fill=tk.BOTH, expand=True)  # Ensure the notebook is displayed properly

# Tab 1: JSON Processing
json_tab = ttk.Frame(notebook)
notebook.add(json_tab, text="JSON Processing")  # Add JSON Processing tab

# Widgets for JSON Processing tab
tk.Label(json_tab, text="JSON Processing Tab").pack(pady=5)

btn_select_files = tk.Button(json_tab, text="Select JSON Files", command=select_files)
btn_select_files.pack(pady=5)

btn_generate_report = tk.Button(json_tab, text="Generate Report", command=generate_report)
btn_generate_report.pack(pady=5)

json_file_listbox_frame = tk.Frame(json_tab)
json_file_listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)

json_file_listbox = tk.Listbox(json_file_listbox_frame, height=15, width=80)
json_scrollbar = tk.Scrollbar(json_file_listbox_frame, orient=tk.VERTICAL, command=json_file_listbox.yview)
json_file_listbox.config(yscrollcommand=json_scrollbar.set)

json_file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
json_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Tab 2: .sor Parsing
sor_tab = ttk.Frame(notebook)
notebook.add(sor_tab, text=".sor Parsing")  # Add .sor Parsing tab

# Variables for .sor parsing tab
sor_input_var = tk.StringVar()
sor_output_var = tk.StringVar()
rbotdr_path_var = tk.StringVar()

tk.Label(sor_tab, text="Input Folder (Containing .sor Files):").pack(pady=5)
tk.Entry(sor_tab, textvariable=sor_input_var, width=60).pack(pady=5)
tk.Button(sor_tab, text="Select Input Folder", command=select_sor_input_folder).pack(pady=5)

tk.Label(sor_tab, text="Output Folder (For Parsed Files):").pack(pady=5)
tk.Entry(sor_tab, textvariable=sor_output_var, width=60).pack(pady=5)
tk.Button(sor_tab, text="Select Output Folder", command=select_sor_output_folder).pack(pady=5)

tk.Label(sor_tab, text="rbOTDR Executable Path:").pack(pady=5)
tk.Entry(sor_tab, textvariable=rbotdr_path_var, width=60).pack(pady=5)
tk.Button(sor_tab, text="Select rbOTDR Executable", command=select_rbotdr_executable).pack(pady=5)

tk.Button(sor_tab, text="Parse .sor Files", command=parse_sor_files).pack(pady=10)

# Run the application
window.mainloop()
