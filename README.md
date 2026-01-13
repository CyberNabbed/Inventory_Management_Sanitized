UNIFIED INVENTORY MANAGEMENT (Sanatized Public Codebase)
=========================
 
OVERVIEW
--------
I wrote this tool to automate the process of updating IT asset inventories. Instead of manually digging through Outlook emails and copy-pasting rows, this program scans a folder of saved email files (.msg), extracts the relevant data, and compiles it all into a single, clean Excel master sheet.
 
It handles two main data sources:
1. Excel Attachments: It pulls rows directly from attached spreadsheets.
2. Email Body Text: It parses "Equipment Loan Agreement" forms (Etrieve) directly from the text of the email.


Initialize UI > Input Selection > Data Extraction > Data Cleaning and Normalization > Master List Verification > Final Export and Summary
 
FEATURES
--------
* Automated Extraction: deeply scans .msg files using 'extract_msg' to grab sender info, dates, and attachments.
* Data Cleaning:
  - Normalizes room numbers (fixing common prefix errors).
  - Standardizes date formats.
  - Fills in missing "Changed By" fields using the sender's name.
* Validation:
  - Enforces Serial Numbers: Any row missing a serial number is automatically dropped to keep data clean.
  - Master List Check: Optionally cross-references a master inventory file and highlights unknown serial numbers in yellow.
* Formatting:
  - Sorts Apple devices to the top.
  - Auto-adjusts column widths.
  - Highlights missing location data or date outliers for easy visual review.
 
REQUIREMENTS
------------
You will need Python 3 installed along with the following libraries:
 
  pip install pandas openpyxl extract-msg easygui
 
(Note: 'tkinter' is also used for the loading screen, but is usually included with standard Python installations.)
 
HOW TO USE
----------
1. Save your target emails as .msg files in a specific folder.
2. Run the script.
3. Follow the on-screen prompts to:
   - Select the Input Folder containing the .msg files.
   - Select the Output Folder where the final report should be saved.
   - (Optional) Select a Master Reference Excel file to check serial numbers against.
4. Review the "final_output.xlsx" file.
   - Yellow cells indicate missing locations, date anomalies, or serial numbers not found in your master list.
 
LOGGING
-------
The tool generates a 'process_log.log' file in the root directory. If a file is skipped or an error occurs, check this log for details.
 
CUSTOMIZATION
-------------
The script contains a configuration section at the top where you can easily adjust:
- Column headers and order.
- Location prefix correction logic.
- Date format preferences.
