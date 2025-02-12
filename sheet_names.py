import pandas as pd

def find_data_sheet(input_file):
    # Open the Excel file
    xls = pd.ExcelFile(input_file, engine='openpyxl')
    
    # Print all sheet names
    print("Available sheets:", xls.sheet_names)
    
    # Try to find a sheet that might contain the raw data
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
        print(f"\nSheet: {sheet_name}")
        print("Columns:", list(df.columns))
        print("Shape:", df.shape)

# Process the file
find_data_sheet('c:/test/2025-01 SunStrong O&M_to_Launch_20250124.xlsx')