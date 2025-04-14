#!/usr/bin/env python3
"""
Teacher Timetable Extraction Tool
---------------------------------
This script extracts teacher timetable data from an Excel file and saves it
in a standardized format to a new Excel file.

Requirements:
    - pandas
    - openpyxl
    - xlsxwriter
    
Install requirements with:
    pip install pandas openpyxl xlsxwriter
"""

import pandas as pd
import os
import sys
from datetime import datetime

def main():
    # Get input and output file paths
    input_file = "04-Ge-Jiao-Shi-Shou-Ke-Shi-Jian-Biao.xlsx"
    output_file = "Formatted_Teacher_Timetables.xlsx"
    
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        sys.exit(1)
    
    print(f"Processing input file: {input_file}")
    print(f"Output will be saved to: {output_file}")
    
    try:
        # Create timestamp for the operation
        start_time = datetime.now()
        
        # Extract and format timetables
        extract_timetables(input_file, output_file)
        
        # Calculate elapsed time
        elapsed_time = datetime.now() - start_time
        
        print(f"\nExtraction completed successfully in {elapsed_time.total_seconds():.2f} seconds.")
        print(f"Formatted timetables saved to: {os.path.abspath(output_file)}")
        
    except Exception as e:
        print(f"Error: An unexpected error occurred: {str(e)}")
        sys.exit(1)

def extract_timetables(input_file, output_file):
    """Extract timetable data from source Excel file and save to a new file."""
    # Read the Excel file to get all sheet names (teacher names)
    print("Reading source Excel file...")
    xls = pd.ExcelFile(input_file)
    teacher_sheets = xls.sheet_names
    
    print(f"Found {len(teacher_sheets)} teacher sheets.")
    
    # Create a new Excel writer object
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    # Define the standard time periods for our output format
    standard_periods = [
        (0, "07:45-08:15"),
        (0, "08:05-08:15"),
        (0, "08:15-08:45"),
        (1, "08:45-09:15"),
        (2, "09:15-09:45"),
        (0, "09:45-10:00"),
        (3, "10:00-10:30"),
        (4, "10:30-11:00"),
        (5, "11:00-11:30"),
        (0, "11:30-11:45"),
        (6, "11:45-12:15"),
        (7, "12:15-12:45"),
        (0, "12:45-13:45"),
        (8, "13:45-14:15"),
        (9, "14:15-14:45"),
        (10, "14:45-15:20"),
        (0, "15:20-15:30"),
        (0, "15:30-16:30")
    ]
    
    # Process each teacher's sheet
    processed_count = 0
    skipped_count = 0
    
    for i, sheet_name in enumerate(teacher_sheets):
        try:
            print(f"Processing sheet {i+1}/{len(teacher_sheets)}: {sheet_name}", end="\r")
            sys.stdout.flush()
            
            # Read the teacher's sheet
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            
            # Skip empty sheets
            if df.empty:
                skipped_count += 1
                continue
                
            # Find period numbers (columns with values 0-10)
            period_col = None
            for col in range(min(3, df.shape[1])):
                numeric_vals = pd.to_numeric(df.iloc[:, col], errors='coerce')
                valid_periods = ((0 <= numeric_vals) & (numeric_vals <= 10)).sum()
                if valid_periods >= 5:  # Found a likely period column
                    period_col = col
                    break
                    
            if period_col is None:
                skipped_count += 1
                continue
                
            # Create a new DataFrame with our standard structure
            timetable = pd.DataFrame(standard_periods, columns=['Period', 'Time'])
            
            # Initialize weekday columns with empty strings
            for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
                timetable[day] = ''
            
            # Map data from source to target
            day_cols = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
            for _, row in df.iterrows():
                # Skip rows without valid period numbers
                try:
                    period = float(row.iloc[period_col])
                    if not (0 <= period <= 10 and period.is_integer()):
                        continue
                    period = int(period)
                except:
                    continue
                    
                # Find matching row in target
                target_rows = timetable[timetable['Period'] == period].index
                if len(target_rows) == 0:
                    continue
                    
                target_idx = target_rows[0]
                
                # Map each day's data
                for i in range(min(5, df.shape[1] - period_col - 1)):
                    src_col = period_col + i + 1
                    if src_col < df.shape[1] and pd.notna(row.iloc[src_col]) and row.iloc[src_col] != '':
                        timetable.at[target_idx, day_cols[i]] = str(row.iloc[src_col]).strip()
            
            # Save to Excel
            safe_name = ''.join(c for c in sheet_name if c not in [':', '\\', '/', '?', '*', '[', ']'])
            if len(safe_name) > 25:
                safe_name = safe_name[:25]
                
            timetable.to_excel(writer, sheet_name=f"Teacher-{safe_name}", index=False)
            
            # Format the worksheet
            workbook = writer.book
            worksheet = writer.sheets[f"Teacher-{safe_name}"]
            
            # Set column widths
            worksheet.set_column('A:A', 8)   # Period
            worksheet.set_column('B:B', 15)  # Time
            worksheet.set_column('C:G', 25)  # Days
            
            processed_count += 1
            
        except Exception as e:
            print(f"\nError processing sheet '{sheet_name}': {str(e)}")
            skipped_count += 1
    
    # Save the Excel file
    writer.close()
    
    print(f"\nSuccessfully processed {processed_count} teacher timetables.")
    if skipped_count > 0:
        print(f"Skipped {skipped_count} sheets due to errors or unsupported formats.")

if __name__ == "__main__":
    main()
