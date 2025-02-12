import pandas as pd

def process_tickets(input_file):
    # Read the Excel file
    df = pd.read_excel(input_file, engine='openpyxl')
    
    # Print exact column names
    print("Actual column names:")
    for col in df.columns:
        print(f"'{col}'")

# Process the tickets
process_tickets('c:/test/2025-01 SunStrong O&M_to_Launch_20250124.xlsx')