import pandas as pd

def read_excel_cell(file_path, row, column):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Select data from the specified cell
        cell_data = df.iloc[row, column]  # Using iloc
        # Alternatively, you can use 'at' accessor
        # cell_data = df.at[row, column]
        
        print(f"Data at cell ({row}, {column}): {cell_data}")
        
        # Return the cell data if needed for further processing
        return cell_data
    
    except Exception as e:
        print("Error:", e)

# Specify the path to the Excel file
excel_file_path = r"C:\Users\jaykumar.bhalani\Downloads\excel.xlsx"

# Specify the row and column to read the cell data
row_number = 1  # Example row number
column_number = 1  # Example column number

# Call the function to read data from the specified cell
read_excel_cell(excel_file_path, row_number, column_number)
