import pandas as pd

def read_excel_file(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Display the DataFrame
        print("Contents of Excel file:")
        print(df)
        
        # Return the DataFrame if needed for further processing
        return df
    
    except Exception as e:
        print("Error:", e)

# Specify the path to the Excel file
excel_file_path = r"C:\Users\jaybh\AppData\Local\Programs\Python\Python311\python.exe"

# Call the function to read the Excel file
read_excel_file(excel_file_path)
