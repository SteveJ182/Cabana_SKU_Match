import pandas as pd

# File paths
file1 = "NEW_Sku_counts.xlsx"  # Path to the first file
file2 = "cleaned_rma.xlsx"  # Path to the second file

# Load the files without headers
df1 = pd.read_excel(file1, header=None)
df2 = pd.read_excel(file2, header=None)

# Assign column names explicitly
df1.columns = ["SKU", "Amount"]
df2.columns = ["SKU", "Amount"]

# Print the number of SKUs in each file
print(f"Number of SKUs in File 1: {len(df1)}")
print(f"Number of SKUs in File 2 (before summing duplicates): {len(df2)}")

# Sum duplicates in File B (df2) by SKU
df2 = df2.groupby("SKU", as_index=False).sum()

# Print the updated number of SKUs in File B
print(f"Number of SKUs in File 2 (after summing duplicates): {len(df2)}")

# Merge all SKUs from both files
merged = pd.merge(df1, df2, on="SKU", how="outer", suffixes=('_file1', '_file2'))

# Define the discrepancy logic
def calculate_discrepancy(row):
    if pd.isna(row["Amount_file1"]) and not pd.isna(row["Amount_file2"]):
        return "Did not receive"
    elif not pd.isna(row["Amount_file1"]) and pd.isna(row["Amount_file2"]):
        return "Not in RMA"
    elif row["Amount_file1"] < row["Amount_file2"]:
        return "Received less"
    elif row["Amount_file1"] > row["Amount_file2"]:
        return "Received additional"
    else:
        return None

# Apply the discrepancy logic to each row
merged["Discrepancy"] = merged.apply(calculate_discrepancy, axis=1)

# Save the result to an Excel file
output_file = "Third_merged_and_discrepancies.xlsx"
merged.to_excel(output_file, index=False)

# Print a summary
print(f"All SKUs combined and saved to {output_file}.")
print(f"Number of discrepancies: {merged['Discrepancy'].notna().sum()}")
