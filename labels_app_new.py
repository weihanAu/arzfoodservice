import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import pandas as pd
import os
import re

def load_file():
    global file_path
    file_path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        label.config(text=f"Loaded: {os.path.basename(file_path)}")

def is_valid(value):
    return value is not None and value != '' and bool(value)

def duplicate_rows_based_on_quantity(file_path):
    # Load the workbook and select the first sheet
    wb = load_workbook(file_path)
    ws = wb.active  # Adjust this if you need a specific sheet

    # Find the last row
    last_row = ws.max_row

    # Loop through rows starting from the last row and working upwards
    for i in range(last_row, 2, -1):  # Starts at row 3
        canPatch = False
        qty_description = ws.cell(row=i, column=10).value
        number_in_parentheses = re.search(r'\((\d+)\)', qty_description)
        if number_in_parentheses:
             number = number_in_parentheses.group(1)
             number = int(number)
             if number:
              canPatch = True
              unit = number
       
        issmallitem = ws.cell(row=i, column=11).value  # Adjust to the correct column for Quantity
        qty = ws.cell(row=i, column=12).value  # Adjust to the correct column for Quantity

        #check unit weight in description
        if  re.search(r'(\d+)g', qty_description):
            issmallitem = True
        if "1kg" in qty_description:
            issmallitem = True
        if "1.5kg" in qty_description:
            issmallitem = True
        if "1.4kg" in qty_description:
            issmallitem = True
        if "1.3kg" in qty_description:
            issmallitem = True
        if "1.2kg" in qty_description:
            issmallitem = True
        if "1.1kg" in qty_description:
            issmallitem = True
        if "0.9kg" in qty_description:
            issmallitem = True
        if "0.8kg" in qty_description:
            issmallitem = True
        if "0.7kg" in qty_description:
            issmallitem = True
        if "0.6kg" in qty_description:
            issmallitem = True
        if "0.5kg" in qty_description:
            issmallitem = True
        if "0.4kg" in qty_description:
            issmallitem = True
        if "0.3kg" in qty_description:
            issmallitem = True
        if "0.2kg" in qty_description:
            issmallitem = True
        if "0.1kg" in qty_description:
            issmallitem = True
        # check if it has strings like '6 x 100 pack'
        if  re.search(r'x\s*\d+', qty_description, re.IGNORECASE):  # Case-insensitive search
            issmallitem = False

        # if canPatch is False, means don't modify the current row
        if canPatch == False:
            unit = 1
            if is_valid(issmallitem):
               if qty == 1:
                ws.cell(row=i, column=12).value=f"{qty}unit" 
               if qty > 1:
                ws.cell(row=i, column=12).value=f"{qty}units" 
               continue
       
        if not isinstance(qty, (int, float)) or not isinstance(unit, (int, float)):
            raise ValueError(f"Error: Row {i} has invalid data types. "
                             f"Quantity: {qty}, Unit: {unit} (Both should be integers).")
        # if unit per box == 1
        if unit == 1:
            if qty ==1:
                ws.cell(row=i, column=12).value = "1 unit"
            while qty > 1:
                ws.cell(row=i, column=12).value = "1 unit"
                ws.insert_rows(i + 1)
                for j in range(1, ws.max_column + 1):
                    ws.cell(row=i + 1, column=j).value = ws.cell(row=i, column=j).value

                ws.cell(row=i + 1, column=12).value = "1 unit"
                ws.cell(row=i + 1, column=13).value = unit
                qty -= 1
            # then stop the codes running    
            continue
        # if unit > 1
        if qty == unit:
            ws.cell(row=i, column=12).value = "1 box"

        while qty > unit:
            ws.cell(row=i, column=12).value = "1 box"
            ws.insert_rows(i + 1)
            for j in range(1, ws.max_column + 1):
                ws.cell(row=i + 1, column=j).value = ws.cell(row=i, column=j).value

            ws.cell(row=i + 1, column=12).value = "1 box"
            ws.cell(row=i + 1, column=13).value = unit
            qty -= unit

        if qty == 1:
               ws.cell(row=i, column=12).value = "1 unit"

        if 1 < qty and qty < unit:
             if is_valid(issmallitem):
                ws.insert_rows(i + 1)  # 插入新行，避免覆盖
                for j in range(1, ws.max_column + 1):  # 复制当前行内容
                        ws.cell(row=i + 1, column=j).value = ws.cell(row=i, column=j).value
                ws.cell(row=i+1, column=12).value = f"{qty} unit" if qty == 1 else f"{qty} units"

             if qty > 1 and not is_valid(issmallitem):
                  while qty > 0 and qty < unit:
                    ws.insert_rows(i + 1)  # 插入新行，避免覆盖
                    for j in range(1, ws.max_column + 1):  # 复制当前行内容
                        ws.cell(row=i + 1, column=j).value = ws.cell(row=i, column=j).value
                    # 更新新行的数据
                    ws.cell(row=i + 1, column=12).value = "1 unit"
                    ws.cell(row=i + 1, column=13).value = unit
                    qty = qty - 1  # 递减 qty
             ws.delete_rows(i)

    base_name, ext = os.path.splitext(file_path)  # Split file name and extension
    new_file_name = f"{base_name}_labels{ext}"   # Add '_labels' to the name
    wb.save(new_file_name)

    # Load the Excel file
    df = pd.read_excel(new_file_name, engine='openpyxl', skiprows=1)
    # reverse the csv file
    df_reversed = df.iloc[::-1]
    # Save the DataFrame to a CSV file
    csv_file = f"{base_name}_labels.csv"  # Change this to your desired output file name
    df_reversed.to_csv(csv_file, index=False)
    
    os.remove(f"{base_name}_labels.xlsx")

def generate_labels():
    if not file_path:
        messagebox.showwarning("Warning", "Please load an Excel file first.")
        return

    try:
        duplicate_rows_based_on_quantity(file_path)
        messagebox.showinfo("Success", f"Labels file created: {os.path.basename(file_path).replace('.xlsx', '_labels.csv')}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Initialize the application
app = tk.Tk()
app.title("Label Generator - ARZ foodservice")

file_path = ""

app.geometry("400x400")

# Create a frame for the UI with padding
frame = tk.Frame(app, padx=20, pady=20)
frame.pack(padx=11, pady=11)

# Load File Button
load_button = tk.Button(app, text="Load Excel File", command=load_file)
load_button.pack(pady=11)

# Label to display loaded file
label = tk.Label(app, text="")
label.pack(pady=11)

# Generate Labels Button
generate_button = tk.Button(app, text="Generate Labels", command=generate_labels)
generate_button.pack(pady=11)

# Text below the Generate Labels button
info_label = tk.Label(app, text=
    "Please don't modify the CSV file downloaded from SAGE. Coz it has the Default structure as below: \n"
    "\n"
    "row 1 is the CSV file title meta.\n"
    "\n"
    "row 2 is the column description.\n"
    "\n"
    "Column G is the quantity that the customer ordered.\n"
    "\n"
    "Column H is the number of items in one box.\n"
    "\n"
    "Column I has the value 'TRUE' is for weighted items.\n"
    "\n",
      wraplength=300)
info_label.pack(pady=5)


app.mainloop()
