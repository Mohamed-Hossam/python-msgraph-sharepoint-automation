🧾 SharePoint List Automation from Excel
This Python script automates the creation, configuration, and population of SharePoint lists using metadata and data defined in an Excel file. It leverages the Microsoft Graph API to interact with SharePoint Online and is ideal for organizations managing structured data across multiple lists with lookup relationships.

📌 Features
🔐 Authentication via Microsoft Identity Platform (MSAL)
📊 Reads metadata from Excel to define list structure, fields, data types, and relationships
🔁 Creates or updates SharePoint lists dynamically
🔗 Supports lookup fields, choice fields, and multi-value lookups
🧹 Deletes existing lists or items before recreation (optional)
📥 Populates lists with initial data from Excel sheets
🧠 Handles data normalization, validation, and error reporting

📁 Excel File Structure
The Excel file (jedco_sharepoint_design.xlsx) must contain the following sheets:

lists_index: Tracks which lists are already created
lists_specs: Defines each list’s fields, types, keys, and relationships
lists_choices: Defines choice values for dropdown fields
One sheet per list: Contains the actual data to be inserted

🛠️ How It Works
Authentication: Uses MSAL to acquire a token for Microsoft Graph.
Site Discovery: Retrieves the SharePoint site ID.
Metadata Parsing: Reads list definitions and relationships from Excel.
List Creation:
Deletes existing lists (if found)
Creates new lists with defined schema
Data Insertion:
Inserts lookup lists first
Then inserts non-lookup lists
Handles lookups, choices, and data types

📦 Requirements
Python 3.7+
Microsoft 365 tenant with SharePoint Online

🚀 Usage
Update the following variables in the script:

client_id = 'YOUR_CLIENT_ID'
client_secret = 'YOUR_CLIENT_SECRET'
tenant_id = 'YOUR_TENANT_ID'
tenant_name = 'yourtenant'
site_name = 'your_site_name'
sharepoint_metadata_file_path = 'your_excel_file.xlsx'


