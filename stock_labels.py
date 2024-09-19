import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
import os

def load_file():
    global file_path
    file_path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        label.config(text=f"Loaded: {os.path.basename(file_path)}")

def duplicate_rows_based_on_quantity(file_path):
    # Load the workbook and select the first sheet
    wb = load_workbook(file_path)
    ws = wb.active  # Adjust this if you need a specific sheet

    # Find the last row
    last_row = ws.max_row

    # Loop through rows starting from the last row and working upwards
    for i in range(last_row, 2, -1):  # Starts at row 3
        qty = ws.cell(row=i, column=7).value  # Adjust to the correct column for Quantity
        unit = ws.cell(row=i, column=8).value  # Get the box unit info
        isWeight = ws.cell(row=i, column=9).value  # Check if this item is a weight item

        if isWeight == '00000000':
            continue

        if not isinstance(qty, (int, float)) or not isinstance(unit, (int, float)):
            raise ValueError(f"Error: Row {i} has invalid data types. "
                             f"Quantity: {qty}, Unit: {unit} (Both should be integers).")

        if qty == unit:
            ws.cell(row=i, column=7).value = "1 box"

        while qty > unit:
            ws.cell(row=i, column=7).value = "1 box"
            ws.insert_rows(i + 1)
            for j in range(1, ws.max_column + 1):
                ws.cell(row=i + 1, column=j).value = ws.cell(row=i, column=j).value

            ws.cell(row=i + 1, column=7).value = "1 box"
            ws.cell(row=i + 1, column=8).value = unit
            qty -= unit

        if 0 < qty < unit:
            ws.cell(row=i, column=7).value = f"{qty} unit" if qty == 1 else f"{qty} units"

    base_name, ext = os.path.splitext(file_path)  # Split file name and extension
    new_file_name = f"{base_name}_labels{ext}"   # Add '_labels' to the name
    wb.save(new_file_name)

def generate_labels():
    if not file_path:
        messagebox.showwarning("Warning", "Please load an Excel file first.")
        return

    try:
        duplicate_rows_based_on_quantity(file_path)
        messagebox.showinfo("Success", f"Labels file created: {os.path.basename(file_path).replace('.xlsx', '_labels.xlsx')}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Initialize the application
app = tk.Tk()
app.title("Label Generator")

file_path = ""

app.geometry("400x300")

# Create a frame for the UI with padding
frame = tk.Frame(app, padx=20, pady=20)
frame.pack(padx=10, pady=10)

# Load File Button
load_button = tk.Button(app, text="Load Excel File", command=load_file)
load_button.pack(pady=10)

# Label to display loaded file
label = tk.Label(app, text="")
label.pack(pady=10)

# Generate Labels Button
generate_button = tk.Button(app, text="Generate Labels", command=generate_labels)
generate_button.pack(pady=10)

app.mainloop()
