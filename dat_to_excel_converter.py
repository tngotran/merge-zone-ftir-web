#!/usr/bin/env python3
"""
DAT to Excel Converter
Converts .dat files to Excel format with proper column names
Handles different data formats automatically
"""

import pandas as pd
import numpy as np
import os
import glob
from pathlib import Path
import re

def install_openpyxl():
    """Install openpyxl if not available"""
    try:
        import openpyxl
        print("✓ openpyxl is available")
        return True
    except ImportError:
        print("Installing openpyxl...")
        import subprocess
        import sys
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
            import openpyxl
            print("✓ openpyxl installed successfully")
            return True
        except Exception as e:
            print(f"✗ Failed to install openpyxl: {e}")
            return False

def detect_column_headers(file_path):
    """
    Detect the column headers and data start line in a DAT file
    Returns: (headers, data_start_line)
    """
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            lines = file.readlines()
        
        for i, line in enumerate(lines):
            line = line.strip()
            
            # Look for lines that contain column headers
            if 'psi(°)' in line and 'Intensity(a.u.)' in line:
                # Format: psi(°) Intensity(a.u.) Sigma_I(a.u.)
                headers = ['psi(°)', 'Intensity(a.u.)', 'Sigma_I(a.u.)']
                return headers, i + 1
            elif 'q(A-1)' in line and 'I(q)' in line:
                # Format: q(A-1) I(q) Sig(q)
                headers = ['q(A-1)', 'I(q)', 'Sig(q)']
                return headers, i + 1
        
        # If no specific headers found, look for the start of numeric data
        for i, line in enumerate(lines):
            line = line.strip()
            if line and not line.startswith('#') and not line.startswith('################################'):
                try:
                    # Try to parse as numbers
                    numbers = [float(x) for x in line.split()]
                    if len(numbers) >= 2:
                        # Default headers based on number of columns
                        if len(numbers) == 3:
                            headers = ['q(A-1)', 'I(q)', 'Sig(q)']
                        elif len(numbers) == 6:
                            headers = ['psi(°)', 'Intensity(a.u.)', 'Sigma_I(a.u.)', 'q(A-1)', 'I(q)', 'Sig(q)']
                        else:
                            headers = [f'Column_{j+1}' for j in range(len(numbers))]
                        return headers, i
                except ValueError:
                    continue
        
        return ['Column_1', 'Column_2', 'Column_3'], 0
        
    except Exception as e:
        print(f"Error detecting headers in {file_path}: {e}")
        return ['Column_1', 'Column_2', 'Column_3'], 0

def parse_dat_file(file_path):
    """
    Parse a .dat file and extract metadata and data with proper column names.
    """
    metadata = {}
    data_lines = []
    
    try:
        # Detect column headers and data start
        column_headers, data_start_line = detect_column_headers(file_path)
        
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            lines = file.readlines()
        
        # Extract metadata from header
        for i, line in enumerate(lines):
            if i >= data_start_line:
                break
                
            line = line.strip()
            if line.startswith('# ') and len(line) > 2:
                # Extract metadata (skip column header lines)
                if not any(header in line for header in ['psi(°)', 'q(A-1)', 'Intensity', 'I(q)']):
                    parts = line[2:].split(None, 1)
                    if len(parts) == 2:
                        key, value = parts
                        metadata[key] = value
        
        # Extract data
        for i in range(data_start_line, len(lines)):
            line = lines[i].strip()
            if not line or line.startswith('#'):
                continue
                
            try:
                numbers = [float(x) for x in line.split()]
                if numbers and len(numbers) >= 2:
                    data_lines.append(numbers)
            except ValueError:
                continue
    
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return None
    
    return {
        'metadata': metadata,
        'data': data_lines,
        'headers': column_headers,
        'file_path': file_path
    }

def convert_dat_to_excel(parsed_data, output_path):
    """
    Convert parsed DAT data to Excel format with proper column names.
    """
    if not parsed_data:
        return False
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            
            # Write metadata to first sheet
            if parsed_data['metadata']:
                metadata_df = pd.DataFrame(list(parsed_data['metadata'].items()), 
                                         columns=['Parameter', 'Value'])
                metadata_df.to_excel(writer, sheet_name='Metadata', index=False)
            
            # Write data to second sheet with proper column names
            if parsed_data['data']:
                data = parsed_data['data']
                headers = parsed_data['headers']
                
                # Determine the number of columns needed
                max_cols = max(len(row) for row in data) if data else 0
                
                if max_cols > 0:
                    # Adjust headers if needed
                    if len(headers) < max_cols:
                        headers.extend([f'Column_{i+1}' for i in range(len(headers), max_cols)])
                    elif len(headers) > max_cols:
                        headers = headers[:max_cols]
                    
                    # Pad rows to same length
                    padded_data = []
                    for row in data:
                        padded_row = row + [None] * (max_cols - len(row))
                        padded_data.append(padded_row[:max_cols])
                    
                    # Create DataFrame with proper column names
                    data_df = pd.DataFrame(padded_data, columns=headers)
                    data_df.to_excel(writer, sheet_name='Data', index=False)
        
        return True
        
    except Exception as e:
        print(f"Error writing Excel file {output_path}: {e}")
        return False

def process_folder(folder_path, base_output_dir):
    """
    Process all .dat files in a folder and convert them to Excel.
    """
    folder_name = os.path.basename(folder_path)
    output_dir = os.path.join(base_output_dir, f"{folder_name}_excel")
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Find all .dat files in the folder
    dat_files = glob.glob(os.path.join(folder_path, "*.dat"))
    
    if not dat_files:
        print(f"No .dat files found in {folder_path}")
        return 0, 0
    
    print(f"Processing {len(dat_files)} .dat files in {folder_name}...")
    
    successful_conversions = 0
    failed_conversions = 0
    
    for dat_file in dat_files:
        # Get filename without extension
        base_name = os.path.splitext(os.path.basename(dat_file))[0]
        excel_file = os.path.join(output_dir, f"{base_name}.xlsx")
        
        # Parse the DAT file
        parsed_data = parse_dat_file(dat_file)
        
        if parsed_data:
            # Convert to Excel
            if convert_dat_to_excel(parsed_data, excel_file):
                successful_conversions += 1
                headers_info = " | ".join(parsed_data['headers'][:3])
                print(f"✓ Converted: {os.path.basename(dat_file)} → {os.path.basename(excel_file)} [{headers_info}]")
            else:
                failed_conversions += 1
                print(f"✗ Failed to convert: {os.path.basename(dat_file)}")
        else:
            failed_conversions += 1
            print(f"✗ Failed to parse: {os.path.basename(dat_file)}")
    
    print(f"Folder {folder_name} - Successful: {successful_conversions}, Failed: {failed_conversions}")
    print(f"Excel files saved to: {output_dir}")
    return successful_conversions, failed_conversions

def process_all_folders(base_dir):
    """
    Process all folders in the base directory.
    """
    print("Starting DAT to Excel conversion process...")
    print("=" * 70)
    
    # Create main output directory
    output_base = os.path.join(base_dir, "excel_conversions")
    os.makedirs(output_base, exist_ok=True)
    
    # Get all subdirectories
    folders = [f for f in os.listdir(base_dir) 
               if os.path.isdir(os.path.join(base_dir, f)) and not f.startswith('.') and f != 'excel_conversions']
    
    if not folders:
        print("No folders found to process.")
        return 0, 0
    
    total_successful = 0
    total_failed = 0
    
    for folder in folders:
        folder_path = os.path.join(base_dir, folder)
        print(f"\n📁 Processing folder: {folder}")
        print("-" * 50)
        
        successful, failed = process_folder(folder_path, output_base)
        total_successful += successful
        total_failed += failed
    
    print("\n" + "=" * 70)
    print("CONVERSION SUMMARY")
    print("=" * 70)
    print(f"Total files processed: {total_successful + total_failed}")
    print(f"Successful conversions: {total_successful}")
    print(f"Failed conversions: {total_failed}")
    if (total_successful + total_failed) > 0:
        success_rate = total_successful/(total_successful + total_failed)*100
        print(f"Success rate: {success_rate:.1f}%")
    
    print(f"\nAll Excel files saved in: {output_base}")
    print("\nColumn formats detected:")
    print("  • Files with psi(°) | Intensity(a.u.) | Sigma_I(a.u.)")
    print("  • Files with q(A-1) | I(q) | Sig(q)")
    
    return total_successful, total_failed

def main():
    """
    Main function to run the conversion process
    """
    # Check and install dependencies
    if not install_openpyxl():
        print("Cannot proceed without openpyxl. Please install it manually.")
        return
    
    # Set the base directory path
    base_directory = "/Users/t.ngo/Desktop/dat2ex"
    
    
    # List all folders in the directory
    folders = [f for f in os.listdir(base_directory) 
               if os.path.isdir(os.path.join(base_directory, f)) and not f.startswith('.')]
    
    print(f"Found {len(folders)} folders to process:")
    total_dat_files = 0
    for folder in folders:
        dat_count = len(glob.glob(os.path.join(base_directory, folder, "*.dat")))
        total_dat_files += dat_count
        print(f"  - {folder}: {dat_count} .dat files")
    
    print(f"\nTotal .dat files to process: {total_dat_files}")
    
    # Run the conversion process
    successful_total, failed_total = process_all_folders(base_directory)
    
    print(f"\n🎉 Conversion process completed!")
    print(f"📊 {successful_total} files converted successfully")
    if failed_total > 0:
        print(f"⚠️  {failed_total} files failed to convert")
    print(f"📁 Check the 'excel_conversions' folder in {base_directory} for your Excel files.")

if __name__ == "__main__":
    main()
