import pandas as pd

# Path to the input Excel file
input_file = input("Enter the path to the Excel file: ")

# Path for the output Excel file
output_file = input("Enter name of cleaned file: ")

try:
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Ensure the first column is processed
    first_column_name = df.columns[0]
    
    # Remove whitespaces from the first column
    df[first_column_name] = df[first_column_name].str.replace(r'\s+', '', regex=True)

    # Save the cleaned data back to an Excel file
    df.to_excel(output_file, index=False)
    print(f"Whitespace removed from the first column. Cleaned file saved as {output_file}")
except Exception as e:
    print(f"An error occurred: {e}")
