import re

def categorize_comment(comment):
    """
    Categorize solar customer comments based on predefined categories.
    
    Args:
        comment (str): The text of the customer comment
    
    Returns:
        str: The matched category or 'No Written Complaint' if no match found
    """
    # Normalize comment to lowercase for easier matching
    comment_lower = comment.lower()
    
    # Define category mapping with keywords and phrases
    categories = {
        'System (Panel) Not Working - with request for cancellation': [
            'cancel', 'cancellation', 'want to cancel', 'cancel the system'
        ],
        'System (Panel) Not Working - with request for payment deferment': [
            'defer payment', 'payment deferment', 'stop payments', 'pause payments'
        ],
        'System (Panel) Not Working - with ticket': [
            'ticket number', 'case number', 'ticket #', 'case #'
        ],
        'System (Panel) Not Working - without ticket': [
            'not working', 'system down', 'panels not generating', 'no power', 
            'system not producing', 'zero production'
        ],
        'System Malfunction - Inverter': [
            'inverter error', 'inverter not working', 'error code', 
            'inverter failure', 'e021', 'state 240', 'red light on inverter'
        ],
        'System Malfunction - Battery/Box': [
            'battery not charging', 'battery issue', 'box problem', 
            'red lights on battery', 'battery shutdown'
        ],
        'System Portal/ App Issue': [
            'app not working', 'cannot view portal', 'production monitoring', 
            'communication issue', 'wifi problem', 'not connecting'
        ],
        'Installation Concerns - Roof Leaks': [
            'roof leak', 'solar-related leak', 'roof damage from installation'
        ],
        'Installation Concerns - Installer Issue': [
            'installer problem', 'installation error', 'improper installation'
        ],
        'Installation Concerns - No Proper Turn Over': [
            'not completed', 'system not hooked up', 'incomplete installation', 
            'not finished', 'permit not obtained'
        ],
        'Installation Concerns - Damage Roof': [
            'roof damage', 'roof collapse', 'roof structural issue'
        ],
        'Billing Statement Issue - Unexplained Charges': [
            'high bill', 'unexpected charges', 'bill much higher', 
            'excessive electricity cost'
        ],
        'Billing Statement Issue - Statement Request': [
            'bill request', 'statement copy', 'billing information'
        ],
        'Billing Statement Issue - Request for Stop Billing': [
            'stop billing', 'cease charges', 'halt payments'
        ],
        'Removel/Re-Installation Request': [
            'remove panels', 'reinstall', 'uninstall', 'roof replacement', 
            'need to remove'
        ],
        'Rebate/ Refund/ Reimbursement Request': [
            'refund', 'rebate', 'reimbursement', 'performance guarantee', 
            'compensation', 'payback'
        ],
        'Account Termination Request': [
            'terminate contract', 'end lease', 'contract termination', 
            'want out of contract'
        ],
        'Account Set-Up Issue': [
            'account setup', 'cannot set up', 'problem creating account', 
            'initialization issue'
        ],
        'Payment Set-Up Issue': [
            'payment setup', 'billing setup', 'cannot set up payment', 
            'payment method problem'
        ],
        'Customer Experience Issues - Delayed Resolution': [
            'no response', 'waiting too long', 'no resolution', 
            'months without fix', 'delayed service'
        ],
        'Customer Experience Issues - Request for Call/ Follow up': [
            'need call back', 'request contact', 'please call', 
            'need to speak with someone'
        ],
        'Regulatory complaint (AG office, state offices, Better Business Bureau, CFPB)': [
            'complaint filed', 'regulatory body', 'attorney general', 
            'better business bureau', 'state office', 'cfpb'
        ],
        'Legal Action from Customer': [
            'legal action', 'lawsuit', 'attorney', 'legal proceeding', 
            'taking legal steps'
        ],
        'Marketing and Promotion': [
            'marketing', 'promotion', 'sales pitch', 'advertisement', 
            'promotional offer'
        ],
        'Pegu Request': [
            'pegu', 'performance', 'performance guarantee'
        ],
        'System Buy-out Procedure': [
            'buy out', 'system purchase', 'purchase option', 'buyout'
        ],
        'Warranty Coverage': [
            'warranty', 'under warranty', 'warranty claim', 'warranty service'
        ]
    }
    
    # Check for matches
    for category, keywords in categories.items():
        if any(keyword in comment_lower for keyword in keywords):
            return category
    
    return 'No Written Complaint'

# Example usage
sample_comments = [
    "My system is not working properly.",
    "I want to cancel my solar contract.",
    "There's a roof leak related to my solar installation.",
    "I need a refund for underperforming panels."
]

for comment in sample_comments:
    print(f"Comment: {comment}")
    print(f"Category: {categorize_comment(comment)}\n")