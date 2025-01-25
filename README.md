# Categorize_Excel_Tickets
Ticket Deduplication and Categorization Script
Dependencies

Python 3.7+
Required Python Libraries:

pandas
openpyxl
re (built-in)



Installation

Install Python from python.org
Install required libraries:
Copypip install pandas openpyxl


Usage

Prepare your Excel file with ticket data
Run the script
When prompted, enter the full path to your Excel file
The script will:

Deduplicate tickets
Categorize tickets based on summary keywords
Save processed tickets to 'deduped_processed_tickets.xlsx'



Features

Removes duplicate tickets based on email, name, and address
Categorizes tickets as:

Red Alert (highest priority)
High Alert
Low Priority


Preserves most recent ticket information

Notes

Requires an Excel file as an input and need to modify the columns if strcture changes.
Expects specific column names and structure
