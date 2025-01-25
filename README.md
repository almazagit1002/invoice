# Invoice Generator and Google Sheets Integration

This program automates the process of generating invoices based on data fetched from a Google Sheet. It creates PDF invoices, updates the Google Sheet status, and integrates company and client details seamlessly.


## Features

- Fetch data from a Google Sheet.
- Generate PDF invoices with a customizable design.
- Automatically mark invoices as "Sent" with the date in the Google Sheet.
- Includes dynamic tax calculation and due date formatting.
- Fully configurable via a `config.json` file.

---

## Prerequisites

Before running the program, ensure the following are installed:

1. **Python** (version 3.8 or above)  
2. **pip** (comes with Python)  
3. **Virtual Environment** (optional but recommended)  

---

## Installation Guide
### 1. Clone the Repository
```bash
git clone https://github.com/your-username/invoice-generator.git
cd invoice-generator
```

### 2. Set Up a Virtual Environment
To keep dependencies isolated, create and activate a virtual environment:
For Windows
```bash
python -m venv venv
.\venv\Scripts\activate  
```
For Mac/Linux:
```bash
source venv/bin/activate
```

### 3. Install Required Dependencies
Run the following command to install all required Python libraries:
```bash
pip install -r requirements.txt
```

## Configuration
### 1. Google Sheets API Setup
Add your Google API credentials to the file keys.json (this file should match the one referenced in config.json).
Ensure the Google Sheet file name and worksheet name are configured in config.json:
```json
"google_sheet": {
  "file": "Invoices",
  "work_sheet": "invoices"}
```
### 2. Customize config.json
Update the config.json file with your company, client, and bank details. An example configuration:
```json
{
    "google_sheets": {
        "key_file": "keys.json",
        "scope": ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    },
    "invoice_directory": "C:/Users/YourUsername/Invoices",
    "company_info": {
        "company_name": "Your Company Name",
        "adress": "Your Address",
        "phone": "Your Phone Number",
        "CVR": "Your CVR Number",
        "email": "Your Email"
    },
    "client_info": {
        "name": "Client Name",
        "adress": "Client Address",
        "CVR": "Client CVR"
    },
    "bank_detial": {
        "bank_name": "Your Bank Name",
        "reg_num": "Your Bank Reg Number",
        "account_number": "Your Account Number"
    },
    "google_sheet": {
        "file": "google sheet file",
        "work_sheet": "google sheet work sheet"
    }
}

```
### 3. Google sheets format
Create a Google Sheet and ensure the following column names are added in the first row of the sheet:
```json 
Place	Case Number	Days	Hours	Day Hour	Hours Worked	Description	Hourly Rate	Subtotal	Moms	Total with Moms	Date	Due Date	Invoice Number	Week	Sent
```
#### Notes:
* Sent: This column should initially contain "N" for unsent invoices. The program updates it to "Y" with the date when the invoice is generated.
* Due Date: Automatically calculated in the Google Sheet using a formula like =TEXT(TODAY(), "DD/MM/YYYY")+15 days or set manually.

## Running the program
Activate the virtual environment:
```bash

.\venv\Scripts\activate
```
Run the main program:
```bash
python main.py
```

## Output
* The program generates PDF invoices and saves them in the directory specified in config.json under invoice_directory.
* The program updates the "Sent" column in the Google Sheet to "Y" and adds the date when the invoice is sent.

## Troubleshooting
* Missing Libraries: Ensure you've run pip install -r requirements.txt in the virtual environment.
* Google API Issues: Ensure your keys.json file is correctly configured and you’ve shared the Google Sheet with the service account email.

## Dependencies
This program uses the following Python libraries:

fpdf (for generating PDF invoices)
gspread (for interacting with Google Sheets)
oauth2client (for authenticating with Google APIs)

## Contributing
Contributions are welcome! Feel free to open issues or submit pull requests for improvements.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Author
Ale Maza
Email : alemazav1002@gmail.com
