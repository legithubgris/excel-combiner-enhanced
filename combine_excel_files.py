#!/usr/bin/env python3
"""
Excel File Combiner Script

This script combines multiple Excel files from a specified folder into a single Excel file.
It reads columns A through C from each Excel file and adds the source filename in column D
only once at the beginning of each file's data group for cleaner output.

Usage:
    python combine_excel_files.py [folder_path]
    
If no folder_path is provided, the script will look for Excel files in the same directory.
"""

import os
import sys
import glob
import pandas as pd
from pathlib import Path
import argparse

def get_excel_files(folder_path, exclude_files=None):
    """
    Get all Excel files from the specified folder.
    
    Args:
        folder_path (str): Path to the folder containing Excel files
        exclude_files (list): List of filenames to exclude
        
    Returns:
        list: Sorted list of Excel file paths
    """
    if exclude_files is None:
        exclude_files = ['combined_excel_files.xlsx', 'test_combined.xlsx']
    
    # Look for both .xlsx and .xls files
    xlsx_pattern = os.path.join(folder_path, "*.xlsx")
    xls_pattern = os.path.join(folder_path, "*.xls")
    
    excel_files = glob.glob(xlsx_pattern) + glob.glob(xls_pattern)
    
    # Filter out excluded files
    filtered_files = []
    for file_path in excel_files:
        filename = os.path.basename(file_path)
        if filename not in exclude_files:
            filtered_files.append(file_path)
    
    # Sort files to ensure consistent order
    filtered_files.sort()
    
    return filtered_files

def read_excel_data(file_path):
    """
    Read Excel file and return data from columns A through C.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        pandas.DataFrame: DataFrame containing the data from columns A-C
    """
    try:
        # Read the Excel file, focusing on columns A, B, C (0, 1, 2)
        df = pd.read_excel(file_path, usecols=[0, 1, 2])
        
        # Get the filename without extension for the source column
        filename = os.path.basename(file_path)
        
        return df, filename
        
    except Exception as e:
        print(f"Error reading file {file_path}: {str(e)}")
        return None, None

def combine_excel_files(folder_path, output_filename="combined_excel_files.xlsx"):
    """
    Combine multiple Excel files into one.
    
    Args:
        folder_path (str): Path to folder containing Excel files
        output_filename (str): Name of the output file
    """
    
    # Get all Excel files in the folder, excluding output files
    exclude_files = [output_filename, 'combined_excel_files.xlsx', 'test_combined.xlsx', 'updated_combined.xlsx', 'final_combined.xlsx', 'final_updated_combined.xlsx']
    excel_files = get_excel_files(folder_path, exclude_files)
    
    if not excel_files:
        print(f"No Excel files found in folder: {folder_path}")
        return
    
    print(f"Found {len(excel_files)} Excel files to combine:")
    for file in excel_files:
        print(f"  - {os.path.basename(file)}")
    
    # Initialize variables for combining data
    combined_data = []
    header_added = False
    
    for file_path in excel_files:
        print(f"\nProcessing: {os.path.basename(file_path)}")
        
        df, source_filename = read_excel_data(file_path)
        
        if df is None:
            continue
            
        # Skip empty files
        if df.empty:
            print(f"  Skipping empty file: {source_filename}")
            continue
        
        # Ensure we have the expected column names or use default ones
        if len(df.columns) >= 3:
            # Rename columns to ensure consistency
            df.columns = ['Filename', 'Transcription', 'Status'] + list(df.columns[3:])
        else:
            print(f"  Warning: File {source_filename} has fewer than 3 columns. Skipping.")
            continue
        
        # Add source filename as column D, but only for the first row
        df['Source_File'] = ''  # Initialize with empty strings
        
        # If this is the first file, include the header
        if not header_added:
            df.iloc[0, df.columns.get_loc('Source_File')] = source_filename  # Add source to first data row
            combined_data.append(df)
            header_added = True
            print(f"  Added header and {len(df)} rows")
        else:
            # For subsequent files, skip the header row (assuming first row is header)
            if len(df) > 1:
                data_rows = df.iloc[1:].copy()  # Skip first row (header) and make a copy
                data_rows['Source_File'] = ''  # Initialize with empty strings
                # Add source filename only to the first row of this batch
                if len(data_rows) > 0:
                    data_rows.iloc[0, data_rows.columns.get_loc('Source_File')] = source_filename
                combined_data.append(data_rows)
                print(f"  Added {len(data_rows)} data rows (skipped header)")
            else:
                print(f"  No data rows to add from {source_filename}")
    
    if not combined_data:
        print("No data to combine!")
        return
    
    # Combine all DataFrames
    final_df = pd.concat(combined_data, ignore_index=True)
    
    # Create output file path
    output_path = os.path.join(folder_path, output_filename)
    
    # Save to Excel
    try:
        final_df.to_excel(output_path, index=False)
        print(f"\nSuccessfully combined {len(excel_files)} files!")
        print(f"Output saved to: {output_path}")
        print(f"Total rows in combined file: {len(final_df)}")
        print(f"Columns: {list(final_df.columns)}")
        
    except Exception as e:
        print(f"Error saving combined file: {str(e)}")

def main():
    """Main function to handle command line arguments and execute the script."""
    
    parser = argparse.ArgumentParser(description='Combine multiple Excel files into one')
    parser.add_argument('folder_path', nargs='?', default='.', 
                       help='Path to folder containing Excel files (default: current directory)')
    parser.add_argument('-o', '--output', default='combined_excel_files.xlsx',
                       help='Output filename (default: combined_excel_files.xlsx)')
    
    args = parser.parse_args()
    
    # Convert to absolute path
    folder_path = os.path.abspath(args.folder_path)
    
    # Check if folder exists
    if not os.path.exists(folder_path):
        print(f"Error: Folder does not exist: {folder_path}")
        sys.exit(1)
    
    if not os.path.isdir(folder_path):
        print(f"Error: Path is not a directory: {folder_path}")
        sys.exit(1)
    
    print(f"Looking for Excel files in: {folder_path}")
    
    # Combine the files
    combine_excel_files(folder_path, args.output)

if __name__ == "__main__":
    main()