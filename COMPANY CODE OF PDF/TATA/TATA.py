import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import fitz  # PyMuPDF
import re
import os

# Function to extract data from a single PDF
def extract_data_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text += page.get_text()

    patterns = {
        'Insured Policy No': re.compile(r'Policy No & Certificate No\s*:\s*(\S+)'),
        'Insured Name': re.compile(r'Insured Name\s*:\s*(.*)'),
        'Customer Phone Number': re.compile(r'Customer contact number\s*:\s*(\S+)'),
        'Customer Email Address': re.compile(r'\s+([\w\.-]+@[\w\.-]+)'),
        'Policy Issued on': re.compile(r'Policy Issuance Date\s*:\s*(\d{2} \w+ \'\d{2})'),
        'Period of Insurance Start': re.compile(r'TP cover period\s*:\s*(\d{2} \w+ \'\d{2}\(00:01Hrs\))'),
        'Period of Insurance End': re.compile(r'TP cover period\s*:\s*\d{2} \w+ \'\d{2}\(00:01Hrs\) to (\d{2} \w+ \'\d{2} \(Midnight\))'),
        'Insured Address': re.compile(r'Address\s*:\s*([\w\s,.-]+\d{6})'),
        'Communication Address': re.compile(r'Address for Communication\s*:\s*([\w\s,.-]+\d{6})'),
        'Vehicle Type': re.compile(r'Vehicle Type\s*:\s*(\w+\s*\w*)'),
        'Make': re.compile(r'Make/Model\s*:\s*([\w\s]+)'),
        'Model': re.compile(r'Make/Model\s*:\s*\w+\s/(\w+)'),
        'Variant': re.compile(r'Variant\s*:\s*(\S+)'),
        'Engine Number': re.compile(r'Engine Number/Battery Number\s*:\s*(\S+)'),
        'Chassis Number': re.compile(r'Chassis number\s*:\s*(\S+)'),
        'Cubic Capacity': re.compile(r'Engine/Battery Capacity \(CC/ KW\)\s*:\s*(\d+)'),
        'Seating Capacity': re.compile(r'Seating Capacity \(including driver\)\s*:\s*(\d)'),
        'Fuel Type': re.compile(r'Fuel Type\s*:\s*(\w+)'),
        'Registration Number': re.compile(r'Registration no\s*:\s*(\S+ \d+ \S+ \d+)'),
        'Date of Registration': re.compile(r'Date of Registration\s*:\s*(\d{2}/\d{2}/\d{4})'),
        'Vehicle IDV': re.compile(r'Insured’s Declared Value\s*:\s*₹ (\d+)'),
        'Manufacturing Year': re.compile(r'Mfg Year\s*:\s*(\d{4})'),
        'Zone': re.compile(r'Zone\s*:\s*(\w+)'),
        'Geographical Area': re.compile(r'Geographical Area\s*:\s*(\w+)'),
        'Body Type': re.compile(r'Body Type\s*:\s*(\w+)'),
        'Policy Expiry Date': re.compile(r'Policy Expiry Date\s*:\s*(\d{2}/\d{2}/\d{4})'),
        'TP Cover Period': re.compile(r'TP cover period\s*:\s*(\d{2} \w+ \'\d{2}\(00:01Hrs\))'),
        "Basic TP Premium": re.compile(r"Basic TP premium\s*₹?\s*([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
        "Basic Own Damage": re.compile(r"Basic Own Damage\s*Premium\s*on\s*Vehicle\s*and\s*non\s*electrical\s*accessories\s*₹?\s*([\d,]+(?:\.\d{2})?)"),
        "Total Own Damage Premium (A)": re.compile(r"Total Own Damage Premium \(A\)\s*:?\s*₹?\s*([\d,]+(?:\.\d{1,2})?)"),
        "Total Liability Premium (B)": re.compile(r"Total Liability Premium \(B\)\s*:?\s*₹?\s*([\d,]+(?:\.\d{1,2})?)"),
        "Net Premium (A+B+C)": re.compile(r"Net Premium \(A\+B\+C\)\s*:?\s*₹?\s*([\d,]+(?:\.\d{1,2})?)"),
        'Liability Premium': re.compile(r"Total Liability Premium \(B\)\s*:?\s*₹?\s*([\d,]+(?:\.\d{1,2})?)"),
        'CGST': re.compile(r'CGST @9 %\s*.\s*(\d+\.\d{2})'),
        'SGST': re.compile(r'SGST @9 %\s*.\s*(\d+\.\d{2})'),
        'IGST': re.compile(r'IGST @18 %\s*.\s*(\d+\.\d{2})'),
        "Total Policy Premium": re.compile(r"Total Policy Premium\s*:?\s*₹?\s*([\d,]+(?:\.\d{1,2})?)"),
        'Nominee Name': re.compile(r'Name of Nominee\s*:\s*(.*)'),  # Updated pattern for Nominee Name
        'Nominee Age': re.compile(r'Age of Nominee\s*:\s*(\d+)'),
        'Nominee Relationship': re.compile(r'Relationship\s*:\s*(.*)'),
        'Nominee Appointee': re.compile(r'Name of Appointee \(if Nominee is minor\)\s*:\s*(.*)'),
        # Updated pattern for Nominee Relationship
        "Previous Policy Number": re.compile(r"Policy Number\s+(\d+)"),
        "Previous Date of expiry": re.compile(r"Date of expiry\s+:\s+(\d{2}/\d{2}/\d{4})"),
        "Type of cover": re.compile(r"Type of cover\s+:\s+(.*)"),
        "Name & address of the Insurer": re.compile(r"Name & address if the Insurer\s+:\s+(.*)")

    }

    extracted_data = {}
    for key, pattern in patterns.items():
        match = pattern.search(text)
        if match:
            extracted_data[key] = match.group(1) if match.lastindex else match.group(0)
        else:
            extracted_data[key] = None

    return extracted_data

# List of PDF files to process
pdf_files = [
    'TATA PDF/6101894270-00.pdf',
    'TATA PDF/6101897717-00.pdf',
    'TATA PDF/6101897726-00.pdf',
    'TATA PDF/6101897729-00.pdf',
    'TATA PDF/6203106795-00.pdf',
    'TATA PDF/t11.pdf',
    'TATA PDF/t12.pdf',
    'TATA PDF/t13.pdf',
    'TATA PDF/6203134035-00.pdf',



]

# Extract data from each PDF and store it
all_data = []
for pdf_file in pdf_files:
    data = extract_data_from_pdf(pdf_file)
    all_data.append(data)

# Create an Excel workbook and sheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Insurance Data"

# Write header
header = list(all_data[0].keys())
ws.append(header)

# Write data
for data in all_data:
    ws.append(list(data.values()))

# Ensure the directory exists
output_directory ='TATA PDF'
os.makedirs(output_directory, exist_ok=True)

# Save the Excel file
excel_path = os.path.join(output_directory, 'Tata_insurance_data.xlsx')
wb.save(excel_path)
print(f"Data extracted and saved to {excel_path}")
