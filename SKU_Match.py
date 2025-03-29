import pandas as pd
import openpyxl
from collections import Counter
import os

# Function to search for UPC in column D and retrieve columns C and D
def find_upc_in_excel(upc_input, df):
    matching_row = df[df.iloc[:, 3] == upc_input]  # Compare input with column D (index 3)
    
    if not matching_row.empty:
        # Pull columns C (index 2) and D (index 3)
        result = matching_row.iloc[:, [2, 3]]  # Retrieve columns C and D
        return result
    else:
        return None

# Function to log found results into the output Excel file
def log_result(upc_input, result, output_file_path):
    try:
        output_df = pd.read_excel(output_file_path, dtype=str)  # Try to load existing output file
    except FileNotFoundError:
        output_df = pd.DataFrame(columns=["SKU", "Scanned"])  # Create new if file doesn't exist

    if result is not None:
        new_entry = pd.DataFrame(result.values, columns=["SKU", "Scanned"])
    else:
        new_entry = pd.DataFrame({"SKU": [None], "Scanned": [upc_input]})

    updated_df = pd.concat([output_df, new_entry], ignore_index=True)
    updated_df.to_excel(output_file_path, index=False)

# Function to log UPCs into the second output Excel file
def log_upc_only(upc_input):
    upc_log_file_path = "scanned_log.xlsx"  # Hardcoded file name for scanned UPC log
    try:
        upc_log_df = pd.read_excel(upc_log_file_path, dtype=str)  # Try to load existing UPC log file
    except FileNotFoundError:
        upc_log_df = pd.DataFrame(columns=["Scanned"])  # Create new if file doesn't exist

    new_upc_entry = pd.DataFrame({"Scanned": [upc_input]})
    updated_upc_df = pd.concat([upc_log_df, new_upc_entry], ignore_index=True)
    updated_upc_df.to_excel(upc_log_file_path, index=False)

# Function to remove whitespaces from the SKU column and save only the cleaned SKU column
def clean_sku_column(input_file, output_file):
    try:
        # Read the Excel file
        df = pd.read_excel(input_file)

        # Ensure the first column (SKU column) is processed
        first_column_name = df.columns[0]  # Assuming the first column is SKU
        
        # Remove whitespaces from the first column
        df[first_column_name] = df[first_column_name].str.replace(r'\s+', '', regex=True)

        # Save only the cleaned SKU column
        cleaned_df = df[[first_column_name]]  # Select only the first column
        cleaned_df.to_excel(output_file, index=False)
    except Exception as e:
        print(f"An error occurred: {e}")

#function to gather all sku's and counts
def count_skus(file_path):
    df = pd.read_excel(file_path, dtype=str)  # Read the Excel file with all columns as text
    first_column_elements = df.iloc[:, 0].dropna().str.strip()  # Select the first column, drop NaNs, and strip spaces
    element_counts = Counter(first_column_elements)
    return element_counts

# File paths
file_path = input("Enter the inventory Excel file to find matching SKU for UPCs: ")
output_file_path = input("Enter the name save the results in an Excel file: ")

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(file_path, dtype=str)  # Read all columns as strings

# Continuous loop to scan and process UPCs
while True:
    upc_input = input("Scan the barcode (UPC) or type 'exit' to quit: ").strip()
    
    if upc_input.lower() == 'exit':
        print("Exiting the program.")
        # Call the function to clean SKU column after exit
        clean_sku_column(output_file_path, "SKUS.xlsx")
        element_counts = count_skus("SKUS.xlsx")
        # Convert the counts to a DataFrame
        counts_df = pd.DataFrame(list(element_counts.items()), columns=['SKU', 'Count'])
        # Ensure the 'SKU' column is treated as text
        counts_df['SKU'] = counts_df['SKU'].astype(str)
        # Sort the DataFrame by SKU
        counts_df.sort_values(by='SKU', ascending=True, inplace=True)
        #output sorted
        counts_df.to_excel("SKUS.xlsx", index=False)
        print("Counts have been saved to SKUS.xlsx")
        os.remove("SKUS.xlsx")
        break

    # Call the function to search for the UPC in the Excel file
    result = find_upc_in_excel(upc_input, df)

    # Log the result into the output Excel file (found or not)
    log_result(upc_input, result, output_file_path)

    # Log the UPC into the hardcoded "scanned_upc_log.xlsx" file
    log_upc_only(upc_input)

    if result is not None:
        print("Matching row found:")
        print(result)
    else:
        print("No matching UPC found.")
