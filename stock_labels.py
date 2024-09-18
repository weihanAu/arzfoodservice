import openpyxl
import os  # Make sure to import the os module

def duplicate_rows_based_on_quantity(file_path, sheet_name):
    # Load the workbook and select the sheet
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]
    
    # Find the last row
    last_row = ws.max_row
    
    # Loop through rows starting from the third row and working upwards
    for i in range(last_row, 2, -1):  # Starts at row 3
        qty = ws.cell(row=i, column=7).value  # Adjust to the correct column for Quantity
        unit = ws.cell(row=i, column=8).value  # Get the box unit info


        # return into next loop if this item is a weight item
        isWeight = ws.cell(row=i, column=9).value # 
        if isWeight:
           print('weight item found, skip it')
           continue
        
        # Check if qty and unit are integers
        if not isinstance(qty, int) or not isinstance(unit, int):
            raise ValueError(f"Error: Row {i} has invalid data types. "
                             f"Quantity: {qty}, Unit: {unit} (Both should be integers).")

        if qty == unit:
            ws.cell(row=i, column=7).value = "1 box"
        
        # Continue processing while the quantity is greater than the unit
        while qty > unit:
            ws.cell(row=i, column=7).value = "1 box"
            
            # Copy the row and insert below the current one
            ws.insert_rows(i + 1)
            for j in range(1, ws.max_column + 1):
                ws.cell(row=i + 1, column=j).value = ws.cell(row=i, column=j).value
            
            ws.cell(row=i + 1, column=7).value = "1 box"
            ws.cell(row=i + 1, column=8).value = unit
            
            # Update the quantity for the next iteration
            qty -= unit
        
        # If the remaining quantity is less than a full unit but more than zero
        if 0 < qty < unit:
            if qty == 1:
                ws.cell(row=i, column=7).value = f"{qty} unit"
            else:
                ws.cell(row=i, column=7).value = f"{qty} units"    
     # Create new file name by adding 'labels' to the original file name
    base_name, ext = os.path.splitext(file_path)  # Split file name and extension
    new_file_name = f"{base_name}_labels{ext}"   # Add '_labels' to the name

    # Save the changes to the workbook
    wb.save(new_file_name)
    print("Rows duplicated and adjusted based on quantity and unit!")
    print("Rows duplicated and adjusted based on quantity and unit!")
    print("Rows duplicated and adjusted based on quantity and unit!")
    print("Rows duplicated and adjusted based on quantity and unit!")

# Usage example:
# just change the file name, then run it.
# just change the file name, then run it.
# just change the file name, then run it.
# just change the file name, then run it.
# just change the file name, then run it.
file_path = 'tmp853A.xlsx'
sheet_name = 'SageReportData1'
duplicate_rows_based_on_quantity(file_path, sheet_name)
