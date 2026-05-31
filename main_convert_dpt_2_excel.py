
import os

# Configure OneDrive path BEFORE importing xlwings
# os.environ['ONEDRIVE_COMMERCIAL_MAC'] = "/Users/t.ngo/Library/CloudStorage/OneDrive-DeakinUniversity"

from unittest import result
import pandas as pd
import xlwings as xw
import glob
from openpyxl import Workbook




# Set the directory containing the .dpt files
#folder_path = "/Users/t.ngo/Downloads/data"
folder_path = "/Users/t.ngo/Desktop/6"

# Step 1
# Convert .dpt files to .xlsx files
final_name = None
for filename in os.listdir(folder_path):
    if filename.endswith('.dpt'):
        file_path = os.path.join(folder_path, filename)
        # Skip empty or near-empty files
        if os.path.getsize(file_path) < 10:
            print(f"Skipping {filename} (empty or corrupt)")
            continue
        # Detect encoding (handle UTF-16 files with BOM)
        with open(file_path, 'rb') as f:
            raw_start = f.read(4)
        if raw_start[:2] in (b'\xff\xfe', b'\xfe\xff'):
            encoding = 'utf-16'
        else:
            encoding = 'utf-8'
        with open(file_path, 'r', encoding=encoding) as f:
            first_line = f.readline()
        sep = ',' if ',' in first_line else '\t' if '\t' in first_line else r'\s+'
        # Use header=None if first line is numeric, header=0 if it's a text header
        try:
            float(first_line.strip().split(',')[0] if ',' in first_line else first_line.strip().split()[0])
            df = pd.read_csv(file_path, sep=sep, engine='python', header=None, names=['Column1', 'Column2'], encoding=encoding)
        except ValueError:
            df = pd.read_csv(file_path, sep=sep, engine='python', header=0, names=['Column1', 'Column2'], encoding=encoding)
        xlsx_filename = filename.rsplit('.', 1)[0] + '.xlsx'
        xlsx_path = os.path.join(folder_path, xlsx_filename)
        df.to_excel(xlsx_path, index=False)
        print(f"Converted {filename} to {xlsx_filename}")
        if final_name is None:
            final_name = filename.rsplit('.', 1)[0]

if final_name is None:
    raise ValueError(f"No .dpt or .xlsx source files found in {folder_path}. Cannot determine output name.")

print(final_name)

# Step 2

# Path to your macro-enabled Excel file
file_path = '/Users/t.ngo/Ny/merge_zone_FTIR/LumosTemplateProtected.xlsm'

# Get or create Excel app instance
if len(xw.apps) > 0:
    app = xw.apps.active
else:
    app = xw.App(visible=True)

# Open the workbook
wb = app.books.open(file_path)
# Run the macro (replace 'MacroName' with your actual macro name)
# Example: wb.macro('MacroName')()
macro_name = 'vbaProject'  # Change this to your actual macro name
wb.macro(macro_name)()


# Step 3
def merge_zone_files(zone_name, output_filename, files=None):
    if files is not None:
        files = sorted(set(files))
    else:
        # Find all xlsx files for the zone (case-insensitive)
        zone_num = zone_name.split()[-1]  # Extract zone number from "ZONE 4"

        # Get all xlsx files and filter case-insensitively
        all_xlsx = glob.glob(os.path.join(folder_path, "*.xlsx"))
        files = []
        for f in all_xlsx:
            fname_lower = os.path.basename(f).lower()
            # Skip already merged files
            if 'merged' in fname_lower:
                continue
            # Check if file contains the zone number (with various formats)
            zone_patterns = [
                f'zone {zone_num}',
                f'zone{zone_num}',
                f'z{zone_num}',
                f'z {zone_num}',
                f'z{zone_num}.',
                f'z {zone_num}.',
                f'zone {zone_num}.',  # For decimal like "Zone 4.0"
                f'zone{zone_num}.',
            ]
            if any(pattern in fname_lower for pattern in zone_patterns):
                files.append(f)

        files = sorted(set(files))
    print(f"Merging {len(files)} files for {zone_name}: {[os.path.basename(f) for f in files]}")
    
    if not files:
        print(f"No files to merge for {zone_name}")
        return

    dfs = [pd.read_excel(f, header=0) for f in files]
    # Find the max number of rows
    max_rows = max(df.shape[0] for df in dfs)
    result_l=[]
    # Prepare merged columns
    merged_cols = []
    for df in dfs:
        # Pad with empty rows if needed
        if df.shape[0] < max_rows:
            df = df.reindex(range(max_rows), fill_value='')

        # print(f"Processing file: {df}")
        merged_cols.append(df.iloc[:,0])
        merged_cols.append(df.iloc[:,1])

        # Copy df.iloc[0] to the first column of the workbook
        wb.sheets[1].range('A1').options(index=False, header=False).value = df.iloc[:, 0].values.reshape(-1, 1)
        # Copy df.iloc[2] to the second column of the workbook
        wb.sheets[1].range('B1').options(index=False, header=False).value = df.iloc[:, 1].values.reshape(-1, 1)

        # Print the result in cell I5
        result = wb.sheets[1]['I5'].value
        result_l.append(result)
        print(f"Result in cell I5: {result}")
        # print("processing complete")
        merged_cols.append([result] + [''] * (max_rows - 1))
        merged_cols.append(['']*max_rows)  # blank column
        
        # Find the row index where df.iloc[:,0] is closest to 1595
        col0_values = pd.to_numeric(df.iloc[:, 0], errors='coerce')
        target = 1595
        closest_idx = (col0_values - target).abs().idxmin()
        value_1172 = float(df.iloc[closest_idx, 1])
        print(f"Row index closest to 1595: {closest_idx}, value in column 1: {df.iloc[closest_idx, 0]}, value in column 2: {value_1172}")

        # value_1172 = df.iloc[1170, 1]
        # print(value_1172)


        target = 2243
        closest_idx = (col0_values - target).abs().idxmin()
        max_local = float(df.iloc[closest_idx, 1])
        print(f"Row index closest to 2243: {closest_idx}, value in column 1: {df.iloc[closest_idx, 0]}, value in column 2: {max_local}")

        # max_local = df.iloc[854, 1]
        # # print(max_local)
        # max_local = df.iloc[855, 1]
        # print(max_local)
        # value_1172 = df.iloc[1173, 1]
        # print(value_1172)
        # value_1172 = df.iloc[1174, 1]
        # print(value_1172)

        # max_local = max (df.iloc[854:856, 1])
        # print(max_local)

        result = ((0.29*value_1172)/((0.29*value_1172)+max_local))*100
        result_l.append(result)

        # Append value_1172, max_local, and result each followed by a blank column
        for val in [value_1172, max_local, result]:
            merged_cols.append([val] + [''] * (max_rows - 1))
        merged_cols.append([''] * max_rows)

    # Combine columns
    merged_cols.append([sum(result_l)/len(result_l) if result_l else '' ] + [''] * (max_rows - 1))    
    merged_df = pd.DataFrame({i: col for i, col in enumerate(merged_cols)})
    merged_df.to_excel(os.path.join(folder_path, output_filename), index=False, header=False)
    print(f"Merged {zone_name} files into {output_filename}")


# If exactly 4 non-merged xlsx files exist, treat them all as Zone 1
all_xlsx = glob.glob(os.path.join(folder_path, "*.xlsx"))
non_merged_xlsx = [f for f in all_xlsx if 'merged' not in os.path.basename(f).lower()]

if len(non_merged_xlsx) == 4:
    print("Exactly 4 xlsx files found — treating all as Zone 1.")
    merge_zone_files('ZONE 1', 'ZONE_1_merged.xlsx', files=non_merged_xlsx)
else:
    # Process zones 1-6, only if files exist
    for zone_num in range(1, 7):
        zone_name = f'ZONE {zone_num}'

        # Get all xlsx files and filter case-insensitively
        all_xlsx = glob.glob(os.path.join(folder_path, "*.xlsx"))

        # Filter for files matching this zone (case-insensitive, with or without space)
        # Matches: "zone 4", "Zone 4", "ZONE 4", "zone4", "Zone4", "ZONE4", "Zone 4.0", etc.
        files = []
        for f in all_xlsx:
            fname_lower = os.path.basename(f).lower()
            # Skip already merged files
            if 'merged' in fname_lower:
                continue
            # Check if file contains the zone number (with various formats)
            zone_patterns = [
                f'zone {zone_num}',
                f'zone{zone_num}',
                f'z{zone_num}',
                f'z {zone_num}',
                f'z{zone_num}.',
                f'z {zone_num}.',
                f'zone {zone_num}.',  # For decimal like "Zone 4.0"
                f'zone{zone_num}.',
            ]
            if any(pattern in fname_lower for pattern in zone_patterns):
                files.append(f)

        if files:
            print(f"Found {len(files)} files for {zone_name}: {[os.path.basename(f) for f in files]}")
            merge_zone_files(zone_name, f'ZONE_{zone_num}_merged.xlsx')
        else:
            print(f"No files found for {zone_name}")

wb.save()
wb.close()


# Step 4
# Merge all merged zone files into one final Excel file, each in a separate sheet
final_excel_path = os.path.join(folder_path, final_name+'_FINAL_OUTPUT.xlsx')

# Check if any merged files exist before creating the writer
merged_files_exist = []
for zone_num in range(1, 7):
    merged_file = os.path.join(folder_path, f'ZONE_{zone_num}_merged.xlsx')
    if os.path.exists(merged_file):
        merged_files_exist.append((zone_num, merged_file))

if merged_files_exist:
    with pd.ExcelWriter(final_excel_path, engine='openpyxl') as writer:
        for zone_num, merged_file in merged_files_exist:
            df = pd.read_excel(merged_file, header=None)
            df.to_excel(writer, sheet_name=f'ZONE {zone_num}', index=False, header=False)
    print(f"Created final file: {final_excel_path}")
else:
    print("No merged zone files found. Skipping Step 4.")

# Delete all .dpt files
for filename in os.listdir(folder_path):
    if filename.endswith('.dpt'):
        os.remove(os.path.join(folder_path, filename))
        print(f"Deleted {filename}")

# # Delete all files with "merged" in the file name
for filename in os.listdir(folder_path):
    if ("merged" in filename or "EXTRACT" in filename) and "OUTPUT" not in filename:
        os.remove(os.path.join(folder_path, filename))
        print(f"Deleted {filename}")

