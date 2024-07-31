import os
import re
import fitz
from openpyxl import Workbook

# Regex patterns for data extraction
patterns = {
    'Policy Number': re.compile(r"Policy / Certificate No\s*:\s*([A-Z0-9]+)"),
    'Insured Name': re.compile(r'Policy\s+Holder\s+Name\s+(.*)'),
    'Intermediary Name': re.compile(r'Intermediary Name\s*:\s*(.*)'),
    'Product Name': re.compile(r'Product\s+Name\s+(.*)'),
    'Policy Start Date': re.compile(r'Policy\s+Start\s+Date\s+(\d{2}/\d{2}/\d{4})'),
    'Policy End Date': re.compile(r'Policy\s+End\s+Date\s+(\d{2}/\d{2}/\d{4})'),
    'Receipt Date': re.compile(r'Receipt\s+Date\s+(\d{2}/\d{2}/\d{4})'),
    'Mobile No': re.compile(r'Proposer\s+Contact\s+Number\s+(\d+)'),
    'Email': re.compile(r'Proposer\s+Email\s+Address\s+([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})'),
    "Address": re.compile(r'Proposer\s+Address\s+([\s\S]*?)\n(\d{6}|\Z)', re.DOTALL),
    'Registration Number': re.compile(r'Registration\s+Number\s+([A-Z0-9]+)'),
    "RTO": re.compile(r'RTO Location\s*\s*(.*)'),
    "Engine Number": re.compile(r'Engine\s+Number\s+(\S+)'),
    "Chassis Number": re.compile(r'Chassis Number\s*\s*(.*)'),
    'Year of Manufacture ': re.compile(r'Year\s+of\s+Manufacture\s+(\d{4})'),
    'make':re.compile(r'Vehicle Make\s*\s*(\w+ \w+)', re.IGNORECASE),
    'Model ': re.compile(r'Vehicle\s+Model\s+(.*)'),
    'Variant ': re.compile(r'Vehicle\s+Variant\s+(.*)'),
    'CC': re.compile(r'Cubic\s+Capacity\s*/\s*Kilo\s+Watt\s*/\s*Gross\s+Vehicle\s+Weight\s*/\s*Horsepower\s+(\d+)'),
    'Seating Capacity including Driver': re.compile(r'Seating\s+Capacity\s+including\s+Driver\s*(\d+)'),
    'Third Party Baisc Premium': re.compile(r'Third Party Baisc Premium\s*(\d+\.\d{2})'),
    'Legal Liability to Driver': re.compile(r'Legal\s+Liability\s+to\s+Driver\s+([\d,.]+)'),
    'Total TP': re.compile(r'TOTAL\s+TP\s+PREMIUM\s+([\d,.]+)'),
    'Total Premium': re.compile(r'TOTAL\s+PREMIUM\s+([\d,.]+)'),
    'GST': re.compile(r'GST\s+([\d,.]+)'),
    'FINAL PREMIUM': re.compile(r'FINAL\s+PREMIUM\s+([\d,.]+)'),

     "Fuel ":  re.compile(r'Fuel\s*\s*(\w+ \(\w+ Kit\))', re.IGNORECASE),


}

# Function to extract text from a PDF
def extract_text_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text += page.get_text()
    return text

# Function to extract details using regex patterns
def extract_details(text, patterns):
    details = {}
    for key, pattern in patterns.items():
        match = pattern.search(text)
        if match:
            details[key] = match.group(1)
        else:
            details[key] = None
    return details

# Function to process all PDFs in a folder
def process_pdfs_in_folder(folder_path):
    extracted_data = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            text = extract_text_from_pdf(pdf_path)
            details = extract_details(text, patterns)
            details['Filename'] = filename  # Add filename as a field
            extracted_data.append(details)
    return extracted_data

# Function to save extracted data to an Excel file
def save_to_excel(data, excel_path):
    wb = Workbook()
    ws = wb.active
    headers = list(patterns.keys()) + ['Filename']
    ws.append(headers)
    for item in data:
        row = [item.get(header, '') for header in headers]
        ws.append(row)
    wb.save(excel_path)

# Main function to run the script
def main():
    folder_path = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\SBI PDF'
    excel_path =  r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\SBI PDF\sbi_pc.xlsx'
    extracted_data = process_pdfs_in_folder(folder_path)
    save_to_excel(extracted_data, excel_path)
    print(f"Data extracted from PDFs in '{folder_path}' and saved to '{excel_path}'.")

if __name__ == "__main__":
    main()
