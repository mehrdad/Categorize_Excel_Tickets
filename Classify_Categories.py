import pandas as pd
import re

VALID_CATEGORIES = [
    'Account Set-Up Issue',
    'Account Termination Request',
    'Billing Statement Issue - Request for Stop Billing',
    'Billing Statement Issue - Statement Request',
    'Billing Statement Issue - Unexplained Charges',
    'Customer Experience Issues - Delayed Resolution',
    'Customer Experience Issues - Request for Call/ Follow up',
    'Installation Concerns  - Damage Roof',
    'Installation Concerns - Installer Issue',
    'Installation Concerns - No Proper Turn Over',
    'Installation Concerns - Roof Leaks',
    'Legal Action from Customer',
    'Marketing and Promotion',
    'No Written Complaint',
    'Payment Set -Up Issue',
    'Pegu Request',
    'Rebate/ Refund/ Reimbursement Request',
    'Regulatory complaint (AG office, state offices, Better Business Bureau, CFPB)',
    'Removel/Re-Installation Request',
    'System (Panel) Not Working - with request for cancellation',
    'System (Panel) Not Working - with request for payment deferment',
    'System (Panel) Not Working - with ticket',
    'System (Panel) Not Working - without ticket',
    'System Buy-out Procedure',
    'System Malfunction -  Inverter',
    'System Malfunction - Battery/Box',
    'System Malfunction - Minor Fix',
    'System Portal/ App Issue',
    'Warranty Coverage'
]

PRIORITY_MAPPING = {
    'Account Set-Up Issue': 'High',
    'Account Termination Request': 'High',
    'Billing Statement Issue - Request for Stop Billing': 'Low',
    'Billing Statement Issue - Statement Request': 'Low',
    'Billing Statement Issue - Unexplained Charges': 'High',
    'Customer Experience Issues - Delayed Resolution': 'Medium',
    'Customer Experience Issues - Request for Call/ Follow up': 'Medium',
    'Installation Concerns  - Damage Roof': 'Medium',
    'Installation Concerns - Installer Issue': 'Medium',
    'Installation Concerns - No Proper Turn Over': 'Medium',
    'Installation Concerns - Roof Leaks': 'High',
    'Legal Action from Customer': 'High',
    'Marketing and Promotion': 'N/A',
    'No Written Complaint': 'Medium',
    'Payment Set -Up Issue': 'Low',
    'Pegu Request': 'Low',
    'Rebate/ Refund/ Reimbursement Request': 'High',
    'Regulatory complaint (AG office, state offices, Better Business Bureau, CFPB)': 'Medium',
    'Removel/Re-Installation Request': 'Medium',
    'System (Panel) Not Working - with request for cancellation': 'High',
    'System (Panel) Not Working - with request for payment deferment': 'Medium',
    'System (Panel) Not Working - with ticket': 'Medium',
    'System (Panel) Not Working - without ticket': 'Low',
    'System Buy-out Procedure': 'High',
    'System Malfunction -  Inverter': 'Medium',
    'System Malfunction - Battery/Box': 'Low',
    'System Malfunction - Minor Fix': 'Low',
    'System Portal/ App Issue': 'Low',
    'Warranty Coverage': 'Low'
}

def categorize_ticket(summary):
    """
    Categorize ticket based on summary text with strict matching to valid categories.
    
    Args:
        summary (str): The ticket summary text
    
    Returns:
        tuple: (category, priority)
    """
    summary_lower = str(summary).lower() if pd.notna(summary) else ''
    
    # Categorization rules with most specific patterns first
    category_rules = [
        ('Account Set-Up Issue', r'account set-?up|cannot access account|login problem'),
        ('Account Termination Request', r'terminate account|cancel service|end contract'),
        ('Billing Statement Issue - Request for Stop Billing', r'stop billing|cease billing'),
        ('Billing Statement Issue - Statement Request', r'bill statement request|copy of bill'),
        ('Billing Statement Issue - Unexplained Charges', r'unexpected charge|unexplained bill|unrecognized charge'),
        ('Customer Experience Issues - Delayed Resolution', r'delayed resolution|waiting too long|unresolved issue'),
        ('Customer Experience Issues - Request for Call/ Follow up', r'call me back|follow ?up|need contact'),
        ('Installation Concerns  - Damage Roof', r'roof damage|damaged roof|installation damage'),
        ('Installation Concerns - Installer Issue', r'installer problem|installation error|poor installation'),
        ('Installation Concerns - No Proper Turn Over', r'no proper turn ?over|incomplete turnover'),
        ('Installation Concerns - Roof Leaks', r'roof leak|leaking roof'),
        ('Legal Action from Customer', r'legal action|lawsuit|legal threat'),
        ('Marketing and Promotion', r'marketing|promotion|promotional'),
        ('No Written Complaint', r'^$|no complaint|no details'),
        ('Payment Set -Up Issue', r'payment set-?up|cannot pay|payment problem'),
        ('Pegu Request', r'pegu request'),
        ('Rebate/ Refund/ Reimbursement Request', r'refund|reimburse|rebate|money back'),
        ('Regulatory complaint (AG office, state offices, Better Business Bureau, CFPB)', r'ag office|better business bureau|cfpb|state office|regulatory complaint'),
        ('Removel/Re-Installation Request', r'remove.*system|re-?install|uninstall'),
        ('System (Panel) Not Working - with request for cancellation', r'not working.*cancellation'),
        ('System (Panel) Not Working - with request for payment deferment', r'not working.*defer'),
        ('System (Panel) Not Working - with ticket', r'not working.*ticket'),
        ('System (Panel) Not Working - without ticket', r'system not working|panel not working'),
        ('System Buy-out Procedure', r'buy-?out|system purchase'),
        ('System Malfunction -  Inverter', r'inverter.*fail|inverter.*error|inverter problem'),
        ('System Malfunction - Battery/Box', r'battery.*issue|box.*problem|battery malfunction'),
        ('System Malfunction - Minor Fix', r'minor fix|small repair|quick repair'),
        ('System Portal/ App Issue', r'portal.*issue|app.*problem|software problem'),
        ('Warranty Coverage', r'warranty|under warranty')
    ]
    
    for category, pattern in category_rules:
        if re.search(pattern, summary_lower):
            return category, PRIORITY_MAPPING.get(category, 'Low')
    
    return 'No Written Complaint', 'Medium'

def process_tickets(input_file):
    """
    Process ticket data from Excel file, categorize tickets, and save results.
    
    Args:
        input_file (str): Path to the input Excel file
    
    Returns:
        pandas.DataFrame: Processed dataframe with categorized tickets
    """
    df = pd.read_excel(
        input_file, 
        sheet_name='2025-01 SunStrong O&M', 
        engine='openpyxl', 
        header=1
    )
    
    # Populate Categories and Priority columns
    df['Categories'], df['Priority'] = zip(*df['Summary'].apply(categorize_ticket))
    
    # Save processed tickets
    output_file = 'processed_tickets.xlsx'
    df.to_excel(output_file, index=False)
    
    return df

# Example usage
if __name__ == "__main__":
    input_file = '2025-01 SunStrong O&M_to_Launch_20250124_byAddress (1).xlsx'  # Replace with your actual file path
    processed_df = process_tickets(input_file)
    print(processed_df[['Summary', 'Categories', 'Priority']].head())