import pandas as pd
import os
import re

def categorize_ticket(summary):
    summary_lower = str(summary).lower()
    # The \b in a regular expression (regex) is a word boundary anchor. It matches the position between a word character 
    # (e.g., letters, numbers, underscores) and a non-word character (e.g., spaces, punctuation, or the start/end of a string). 
    red_alert_keywords = [
        r'\bcfpb\b', r'\battorney general\b', r'\bag\b', r'\bclass action\b', r'\blawsuit\b'
    ]
    
    high_alert_keywords = [
        r'\bnews\b', r'\bmedia\b', r'\blawyer\b', r'\battorney\b', 
        r'\bbbb\b', r'\bbetter business bureau\b', r'\billinois shines\b', r'\bil shines\b'
    ]
    
    for keyword in red_alert_keywords:
        if re.search(keyword, summary_lower):
            return 'Red Alert', 3
    
    for keyword in high_alert_keywords:
        if re.search(keyword, summary_lower):
            return 'High Alert', 2
    
    return 'Low Priority', 1

def process_tickets(input_file):
    # Read the Excel file with specified headers
    df = pd.read_excel(
        input_file, 
        sheet_name='2025-01 SunStrong O&M', 
        engine='openpyxl', 
        header=1
    )
    
    # Select and filter specific columns
    selected_columns = [
        'Categories', 'Priority', '#', 'TicketId', 'Title', 'Summary',
        'CreatedDate', 'UpdatedDate', 'GroupName', 'FormName', 'TicketStatus', 'RequesterName',
        'RequesterEmail', 'RequesterPersonId', 'SunStrong- Account #',
        'Your Name', 'Address of Solar System', 'Was it PTO?', 'Date System was Installed', 'Assignment'
    ]
    
    # Ensure only the selected columns are retained
    processed_tickets = df[selected_columns].copy()
    
    # Normalize email and name for deduplication
    processed_tickets['NormalizedEmail'] = processed_tickets['RequesterEmail'].str.strip().str.lower()
    processed_tickets['NormalizedName'] = processed_tickets['RequesterName'].str.strip().str.lower()
    processed_tickets['NormalizedAddress'] = processed_tickets['Address of Solar System'].str.strip().str.lower()
    
    # Deduplicate with multiple criteria
    processed_tickets = processed_tickets.sort_values('CreatedDate', ascending=False).drop_duplicates(
        subset=['NormalizedEmail', 'NormalizedName', 'NormalizedAddress'], 
        keep='first'
    )
    
    # Remove temporary normalization columns
    processed_tickets = processed_tickets.drop(columns=['NormalizedEmail', 'NormalizedName', 'NormalizedAddress'])
    
    # Add Category and Priority columns
    processed_tickets['Category'], processed_tickets['Priority'] = zip(
        *processed_tickets['Summary'].apply(categorize_ticket)
    )
    
    # Save processed tickets
    output_file = 'deduped_processed_tickets.xlsx'
    processed_tickets.to_excel(output_file, index=False)
    
    return processed_tickets

def main():
    # Prompt for input file
    input_file = input("Enter the full path to the Excel file: ").strip()
    
    # Validate file exists
    if not os.path.exists(input_file):
        print(f"Error: File {input_file} does not exist.")
        return
    
    # Process tickets
    result = process_tickets(input_file)
    print(f"Processed tickets saved to 'deduped_processed_tickets.xlsx'. Total records: {len(result)}")

if __name__ == "__main__":
    main()