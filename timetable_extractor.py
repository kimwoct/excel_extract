#!/usr/bin/env python3

import sys
import pandas as pd
import numpy as np

def main():
    # Get input and output filenames from command line arguments or use defaults
    input_file = sys.argv[1] if len(sys.argv) > 1 else "04-Ge-Jiao-Shi-Shou-Ke-Shi-Jian-Biao.xlsx"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "Processed_Teacher_Timetables.xlsx"

    print(f"Processing input file: {input_file}")
    print(f"Output will be saved to: {output_file}")

    try:
        # Read all sheets
        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            for sheet_name in sheet_names:
                try:
                    df = pd.read_excel(input_file, sheet_name=sheet_name)
                    if df.empty:
                        continue

                    # Find the header row (usually the row with '課節' or 'Period')
                    header_row = df[df.iloc[:,0].astype(str).str.contains('課節|Period', na=False)].index
                    if len(header_row) == 0:
                        continue
                    header_row = header_row[0]
                    
                    # Store the original data before setting headers
                    original_data = df.copy()
                    
                    # Set column headers
                    df.columns = df.iloc[header_row]
                    df = df.iloc[header_row+1:]

                    # Clean up columns
                    df = df.reset_index(drop=True)
                    df = df[['課節', '時段', '星期一', '星期二', '星期三', '星期四', '星期五']]
                    df.columns = ['Period', 'Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

                    # Function to combine rows for each timeslot
                    def combine_rows(main_row, rows_above):
                        result = {}
                        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
                        
                        for day in days:
                            values = []
                            # Add values from rows above if they exist and aren't empty
                            for row in rows_above:
                                if pd.notna(row.get(day)) and str(row.get(day)).strip():
                                    values.append(str(row.get(day)).strip())
                            # Add main row value if it exists and isn't empty
                            if pd.notna(main_row.get(day)) and str(main_row.get(day)).strip():
                                values.append(str(main_row.get(day)).strip())
                            
                            result[day] = '-'.join(values) if values else ''
                            
                        result['Period'] = main_row.get('Period', '')
                        result['Time'] = main_row.get('Time', '')
                        return result

                    # Process the data with additional rows
                    processed_rows = []
                    i = 0
                    while i < len(df):
                        main_row = df.iloc[i].to_dict()
                        
                        # Get the next two rows if they exist and have the same time slot
                        additional_rows = []
                        for j in [1, 2]:
                            if i + j < len(df) and df.iloc[i + j]['Time'] == main_row['Time']:
                                additional_rows.append(df.iloc[i + j].to_dict())
                                i += 1
                            
                        combined_row = combine_rows(main_row, additional_rows)
                        processed_rows.append(combined_row)
                        i += 1

                    # Create new DataFrame from processed rows
                    grouped = pd.DataFrame(processed_rows)
                    
                    # Reorder columns
                    grouped = grouped[['Period', 'Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']]

                    # Save to Excel with adjusted row heights
                    safe_name = ''.join(c for c in sheet_name if c not in [':', '\\', '/', '?', '*', '[', ']'])
                    if len(safe_name) > 25:
                        safe_name = safe_name[:25]
                    
                    # Write to Excel with formatting
                    grouped.to_excel(writer, sheet_name=safe_name, index=False)
                    
                    # Get the worksheet object
                    worksheet = writer.sheets[safe_name]
                    
                    # Set row height for all rows to accommodate multiple lines
                    for row in range(len(grouped) + 1):  # +1 for header row
                        worksheet.set_row(row, 45)  # Set row height to 45 points
                    
                    # Set column widths
                    worksheet.set_column('A:B', 15)  # Period and Time columns
                    worksheet.set_column('C:G', 25)  # Day columns

                except Exception as e:
                    print(f"Error processing sheet '{sheet_name}': {str(e)}")

    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
