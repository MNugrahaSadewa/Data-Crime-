import pandas as pd

# Load the dataset from the Excel file
file_path = 'F:/Tel u/Data Sains/data Crime/crime dataset separate.xlsx'
df = pd.read_excel(file_path, sheet_name='crime dataset separate')

# Get the total number of records in the dataset
total_records = len(df)
print(f"Total number of records in the crime dataset: {total_records}")
