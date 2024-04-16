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

def insert_data_to_excel(file_path, new_data):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Append new data to the DataFrame
        df = df.append(new_data, ignore_index=True)
        
        # Write the updated DataFrame back to the Excel file
        df.to_excel(file_path, index=False)
        
        print("New data inserted successfully.")
    
    except Exception as e:
        print("Error:", e)

# Specify the path to the Excel file
excel_file_path = r"C:\Users\jaykumar.bhalani\Downloads\excel.xlsx"

# Call the function to read the Excel file
existing_data = read_excel_file(excel_file_path)

# Define new data to insert into the Excel file
new_data = pd.DataFrame({
    'ID': ['9'],
    'NAME': ['Ellyse Perry'],
    'DEPARTMENT': [Computer],
})

# Call the function to insert new data into the Excel file
insert_data_to_excel(excel_file_path, new_data)

# Call the function to read the Excel file after insertion
read_excel_file(excel_file_path)
