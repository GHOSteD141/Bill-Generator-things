import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def read_excel(file_path):
    # Read Excel/CSV data
    df = pd.read_excel(file_path)
    return df

def generate_pdf_bill(data, output_path):
    # Use reportlab to generate PDF
    c = canvas.Canvas(output_path, pagesize=A4)
    # ...customize bill layout...
    c.drawString(100, 800, f"Bill for {data['Name']}")
    # ...more customization...
    c.save()

def read_google_sheet(json_keyfile, sheet_name):
    """
    Reads data from a Google Sheet linked to Google Forms.
    :param json_keyfile: Path to your Google service account JSON key.
    :param sheet_name: Name of the Google Sheet.
    :return: pandas DataFrame with the sheet data.
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    return df

def generate_bills_from_excel(excel_path, output_dir):
    """
    Reads Excel data and generates a PDF bill for each entry.
    :param excel_path: Path to the Excel file.
    :param output_dir: Directory to save generated bills.
    """
    df = read_excel(excel_path)
    for idx, row in df.iterrows():
        # Customize filename as needed, e.g., using Name or ID
        filename = f"{row.get('Name', 'bill')}_{idx+1}.pdf"
        output_path = f"{output_dir}/{filename}"
        generate_pdf_bill(row, output_path)

# ...existing code for DOCX, HTML, etc...
