import pandas as pd

def categorize_ticket(summary):
    summary_lower = str(summary).lower()
    
    red_alert_keywords = [
        'cfpb', 'attorney general', 'ag', 'class action', 'lawsuit'
    ]
    
    high_alert_keywords = [
        'news', 'media', 'lawyer', 'attorney', 'bbb', 
        'better business bureau', 'illinois shines', 'il shines'
    ]
    
    for keyword in red_alert_keywords:
        if keyword in summary_lower:
            return 'Red Alert', 3
    
    for keyword in high_alert_keywords:
        if keyword in summary_lower:
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
        'TicketId', 'Title', 'Summary', 
        'RequesterName', 'RequesterEmail', 
        'Address of Solar System', 'Assignment'
    ]
    
    # Filter for Launch tickets and select specific columns
    launch_tickets = df[df['Assignment'] == 'Launch'][selected_columns].copy()
    
    # Add Category and Priority columns
    launch_tickets['Category'], launch_tickets['Priority'] = zip(
        *launch_tickets['Summary'].apply(categorize_ticket)
    )
    
    # Save processed tickets
    output_file = 'processed_tickets.xlsx'
    launch_tickets.to_excel(output_file, index=False)
    
    return launch_tickets

# Process the tickets
result = process_tickets('c:/test/2025-01 SunStrong O&M_to_Launch_20250124.xlsx')
print(result)