from openpyxl import load_workbook

def update_excel_cell(file_path, sheet_name, row, column, new_value):
    try:
        # Load the workbook
        wb = load_workbook(filename=file_path)
        
        # Select the active worksheet
        ws = wb[sheet_name]
        
        # Update the value in the specified cell
        ws.cell(row=row, column=column, value=new_value)
        
        # Save the workbook
        wb.save(file_path)
        
        print(f"Data in cell ({row}, {column}) updated successfully.")
    
    except Exception as e:
        print("Error updating data in Excel file:", e)

# Specify the path to the Excel file
excel_file_path = r"D:\task6\excel.xlsx"

# Specify the name of the worksheet
sheet_name = "Sheet1"  # Example sheet name

# Specify the row and column to update the cell data
row_number = 3  # Example row number
column_number = 2  # Example column number

# Specify the new value to update the cell with
new_cell_value = "JAHAN"  # Example new value

# Call the function to update data in the specified cell
update_excel_cell(excel_file_path, sheet_name, row_number, column_number, new_cell_value)

