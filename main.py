import pandas as pd

# Load the Excel file
file_path = 'sample_file.xlsx'

# Read all sheets
sima_df = pd.read_excel(file_path, sheet_name='Quantity on Order - SIMA System', engine='openpyxl')
allocation_df = pd.read_excel(file_path, sheet_name='Allocation File', engine='openpyxl')
orders_df = pd.read_excel(file_path, sheet_name='Orders-SIMA System', engine='openpyxl')

# Display basic shapes for confirmation
print("✅ Loaded Sheets")
print(f"Quantity on Order - SIMA: {sima_df.shape}")
print(f"Allocation File: {allocation_df.shape}")
print(f"Orders-SIMA System: {orders_df.shape}")
