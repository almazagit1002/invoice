import json
import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fpdf import FPDF
from datetime import datetime

# Load configuration file
def load_config(config_path="config.json"):
    with open(config_path, "r") as config_file:
        return json.load(config_file)

config = load_config()

# Google Sheets Setup
def get_google_sheet_data(sheet_name, worksheet_name):
    scope = config["google_sheets"]["scope"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(config["google_sheets"]["key_file"], scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name)
    worksheet = sheet.worksheet(worksheet_name)
    return worksheet

# Custom Invoice Class
class CustomInvoice(FPDF):
    def __init__(self, data):
        super().__init__()  # Initialize the FPDF class properly
        self.data = data

    def header(self):
        # Add company name from config
        company_info = config["company_info"]
        client_info = config["client_info"]

        self.set_y(10)  # Align to the top of the page
        self.set_font("Arial", "B", 20)
        self.cell(0, 10, company_info["company_name"], ln=True, align="R")

        self.set_font("Arial", "B", 14)
        self.cell(0, 5, client_info["name"], ln=True, align="L")
        self.set_font("Arial", "", 10)
        self.cell(0, 5, client_info["adress"], ln=True, align="L")
        self.cell(0, 5, client_info["CVR"], ln=True, align="L")

    def footer(self):
        bank_details = config["bank_detial"]
        company_info = config["company_info"]

        self.set_y(-110)
        self.set_font("Arial", "I", 9)
        self.multi_cell(0, 10, 
                        f"Betalingsbetingelser: Netto 15 dage - Forfaldsdato: {self.data['Due Date']}\n"
                        f"Beløbet indbetales på bankkonto:\n"
                        f"{bank_details['bank_name']} / {bank_details['reg_num']} / {bank_details['account_number']}\n"
                        f"Fakturanr. angives ved bankoverførsel.\n"
                        f"Ved betaling efter forfald tilskrives der renter på 0,93%, pr. måned, samt et gebyr på 100,00 DKK.",
                        align="L")

        self.set_y(-30)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())

        self.set_y(-25)
        self.set_font("Arial", "I", 9)
        self.cell(0, 10, 
                  f"{company_info['company_name']} / {company_info['adress']} / {company_info['phone']} / {company_info['email']}", 
                  ln=True, align="C")


# Generate Invoice PDF
def create_invoice(data, output_file):
    pdf = CustomInvoice(data)
    pdf.add_page()
    
    pdf.set_font("Arial", "", 10)
    dato_width = pdf.get_string_width("Dato: ")
    pdf.cell(dato_width, 10, "Dato: ", ln=False, align="L")
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 10, data['Date'], ln=False, align="L")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, f"Fakturanr. {data['Invoice Number']}", ln=True, align="R")

    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, f"Week {data['Week']} - {data['Place']} - {data['Description']} - Rope Access - Case Number: {data['Case Number']}", ln=True)
    pdf.ln(10)

    pdf.set_fill_color(200, 200, 200)
    pdf.cell(60, 10, "Beskrivelse", border=1, fill=True)
    pdf.cell(40, 10, "Antal", border=1, align="C", fill=True)
    pdf.cell(40, 10, "Enhedspris", border=1, align="C", fill=True)
    pdf.cell(50, 10, "Pris", border=1, align="C", fill=True)
    pdf.ln(10)

    pdf.set_fill_color(255, 255, 255)
    pdf.cell(60, 10, "Rope Access", border=1)
    pdf.cell(40, 10, f"{data['Hours Worked']:.2f} timer", border=1, align="C")
    pdf.cell(40, 10, f"{data['Hourly Rate']:.2f}", border=1, align="C")
    total_price = data['Hours Worked'] * data['Hourly Rate']
    pdf.cell(50, 10, f"{total_price:.2f}", border=1, align="C")
    pdf.ln(10)

    tax = total_price * 0.25
    grand_total = total_price + tax

    pdf.cell(0, 10, f"Subtotal: {total_price:.2f} DKK", ln=True, align="R")
    pdf.cell(0, 10, f"Moms (25,00%): {tax:.2f} DKK", ln=True, align="R")
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 10, f"Total: {grand_total:.2f} DKK", ln=True, align="R")

    pdf.output(output_file)
    print(f"Invoice saved as {output_file}")

# Main Function
if __name__ == "__main__":
    sheet_name = config["google_sheet"]["file"]
    worksheet_name = config["google_sheet"]["work_sheet"]
    worksheet = get_google_sheet_data(sheet_name, worksheet_name)
    records = worksheet.get_all_records()

    for idx, record in enumerate(records):
        if record["Sent"] == "N":
            date_str = datetime.now().strftime("%Y%m%d")
            output_filename = os.path.join(config["invoice_directory"], f"Faktura-{record['Invoice Number']}.pdf")
            create_invoice(record, output_filename)

            worksheet.update_cell(idx + 2, list(record.keys()).index("Sent") + 1, "Y")
            print(f"Updated 'Sent' status for Invoice Number: {record['Invoice Number']}")
