import pandas as pd
import os
import win32com.client
from datetime import datetime
import shutil

# --- CONFIGURATION ---
# Use absolute paths for Excel COM automation
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, 'Header_Template.xlsx')
SOURCE_DATA = os.path.join(BASE_DIR, 'Sodexo Accounts 7-1-2025 to 1-31-2026.xlsx')
OUTPUT_DIR = os.path.join(BASE_DIR, 'Split_Categories')
FILTER_COLUMN = 'Customer Name'

# --- DATA LOADING (Pandas is still faster for reading) ---
def find_header_row_with_keyword(sheet_name, keywords=('customer', 'name'), max_rows=25):
    try:
        tmp = pd.read_excel(SOURCE_DATA, sheet_name=sheet_name, header=None, nrows=max_rows)
        for i in range(len(tmp)):
            row_str = tmp.iloc[i].astype(str).str.replace('\n', ' ').str.replace('\r', ' ').str.lower().str.cat(sep=' ')
            if all(k in row_str for k in keywords):
                return i
    except Exception as e:
        print(f"Warning: Could not scan headers for {sheet_name}: {e}")
    return None

def normalize_columns(df):
    cols = df.columns.astype(str)
    cols = cols.str.replace('\r', ' ').str.replace('\n', ' ').str.strip().str.replace(r'\s+', ' ', regex=True)
    df.columns = cols
    
    if FILTER_COLUMN not in df.columns:
        candidates = [c for c in df.columns if 'customer' in c.lower() and ('name' in c.lower() or 'number' in c.lower() or '#' in c)]
        if candidates:
            df = df.rename(columns={candidates[0]: FILTER_COLUMN})
    return df

def load_source_sheet(sheet_name):
    print(f"Detecting headers for source sheet '{sheet_name}'...")
    header_idx = find_header_row_with_keyword(sheet_name)
    if header_idx is None:
        header_idx = 0 
        print(f"  Header check failed, using fallback row {header_idx + 1}")
    else:
        print(f"  Found headers at row {header_idx + 1}")
    
    df = pd.read_excel(SOURCE_DATA, sheet_name=sheet_name, header=header_idx)
    df = normalize_columns(df)
    df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed')]
    return df

print(f"[{datetime.now().strftime('%H:%M:%S')}] Loading master data...")
df_produce = load_source_sheet('Produce')
df_broadline = load_source_sheet('Broadline')

df_produce[FILTER_COLUMN] = df_produce[FILTER_COLUMN].astype(str).str.strip()
if FILTER_COLUMN in df_broadline.columns:
    df_broadline[FILTER_COLUMN] = df_broadline[FILTER_COLUMN].astype(str).str.strip()

# Get unique categories from both sheets
cats_p = set(df_produce[FILTER_COLUMN].dropna().unique()) if FILTER_COLUMN in df_produce.columns else set()
cats_b = set(df_broadline[FILTER_COLUMN].dropna().unique()) if FILTER_COLUMN in df_broadline.columns else set()

# Combine and sort, removing 'nan' strings
categories = sorted(list((cats_p | cats_b) - {'nan', 'NaN', 'None'}))

# --- EXCEL AUTOMATION ---
def assemble_files_batch(categories, df_p, df_b):
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # Start Excel instance
    print("Starting Excel instance...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False # Speed up execution

    try:
        count = 0
        total = len(categories)
        
        for cat in categories:
            try:
                clean_name = str(cat).replace("/", "-").strip()
                cat_folder = os.path.join(OUTPUT_DIR, clean_name)
                if not os.path.exists(cat_folder):
                    os.makedirs(cat_folder)
                
                target_path = os.path.join(cat_folder, f"{clean_name}_Report.xlsx")
                
                # Filter data
                sub_p = df_p[df_p[FILTER_COLUMN] == cat]
                sub_b = df_b[df_b[FILTER_COLUMN] == cat] if FILTER_COLUMN in df_b.columns else pd.DataFrame()
                
                if sub_p.empty and sub_b.empty:
                    continue

                # Copy template first (fast file system copy)
                shutil.copy2(TEMPLATE_FILE, target_path)
                
                # Open the copied file in Excel
                wb = excel.Workbooks.Open(target_path)
                
                # Update Information Sheet
                try:
                    ws_info = wb.Sheets('Information')
                    ws_info.Range('D5').Value = cat
                    ws_info.Range('D9').Value = datetime.now().strftime('%m/%d/%Y')
                    
                    if not sub_b.empty and 'Invoice Date' in sub_b.columns:
                        b_dates = pd.to_datetime(sub_b['Invoice Date'], errors='coerce').dropna()
                        if not b_dates.empty:
                            ws_info.Range('D7').Value = b_dates.min().strftime('%m/%d/%Y')
                            ws_info.Range('D8').Value = b_dates.max().strftime('%m/%d/%Y')
                except Exception as e:
                    print(f"  Warning (Info Sheet): {e}")

                # Update Data Sheets
                mappings = {'Produce': sub_p, 'Broadline': sub_b}
                for sheet_name, data in mappings.items():
                    try:
                        ws = wb.Sheets(sheet_name)
                        if not data.empty:
                            # Convert to list of lists for fast block writing
                            # Ensure dates are strings to avoid Excel auto-formatting issues
                            # Excel expects 1-based indexing for Range
                            
                            # Pre-process data for Excel
                            # Fill NaNs with empty string
                            data = data.fillna('')
                            
                            # Convert Timestamps to strings
                            for col in data.select_dtypes(include=['datetime', 'datetimetz']).columns:
                                data[col] = data[col].dt.strftime('%m/%d/%Y')
                                
                            values = data.values.tolist()
                            
                            if values:
                                start_row = 2
                                start_col = 1
                                end_row = start_row + len(values) - 1
                                end_col = start_col + len(values[0]) - 1
                                
                                # Define range
                                # Use Cells(row, col) to get Range object
                                rng = ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col))
                                rng.Value = values

                                # Clear any remaining cells
                                if sheet_name == 'Produce':
                                    ws.Rows(2).RowHeight = 30  # Adjust number as needed

                                
                    except Exception as e:
                        # Sheet might not exist
                        pass

                wb.Save()
                wb.Close()
                
                count += 1
                if count % 5 == 0:
                    print(f"  Processed {count}/{total}: {cat}")
                else:
                    print(f"  Success: {cat}")

            except Exception as e:
                print(f"Failed {cat}: {e}")

    finally:
        print("Closing Excel...")
        excel.Quit()
        # Kill excel process if stuck? No, just Quit should work.

print(f"Processing {len(categories)} categories with Excel Automation...")
assemble_files_batch(categories, df_produce, df_broadline)

print("-" * 30)
print(f"Finished! Files are in {OUTPUT_DIR}")
