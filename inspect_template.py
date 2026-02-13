
import pandas as pd
import os

TEMPLATE_FILE = 'Header_Template.xlsx'

try:
    print(f"Inspecting {TEMPLATE_FILE}...")
    for sheet in ['Produce', 'Broadline']:
        print(f"\n--- Sheet: {sheet} ---")
        # Read first 5 rows with no header to see the actual structure
        df = pd.read_excel(TEMPLATE_FILE, sheet_name=sheet, header=None, nrows=5)
        print("First 5 rows (raw):")
        print(df)
        
        # Try to detect header row
        for i, row in df.iterrows():
            row_vals = [str(x).lower() for x in row if pd.notna(x)]
            if any('source' in x or 'date' in x or 'customer' in x for x in row_vals):
                print(f"Potential header found at row index {i}: {list(row)}")
        
except Exception as e:
    print(f"Error: {e}")
