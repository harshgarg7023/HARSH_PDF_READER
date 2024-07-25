import os
import fitz  # PyMuPDF
import re
import pandas as pd

# Function to extract data from a single PDF
def extract_data_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text += page.get_text()

    patterns = {
        'Policy Number': re.compile(r'Policy Number\s*:\s*(\S+)'),
        'Covernote No': re.compile(r'Covernote No\s*:\s*(\S+)'),
        'Insured Name': re.compile(r'Insured Name\s*:\s*(.*?)(?=\n|$)'),
        'Period of Insurance Start': re.compile(r'From\s*00:00 Hrs on\s*(\d{2}-\w+-\d{4})'),
        'Period of Insurance End': re.compile(r'to Midnight of\s*(\d{2}-\w+-\d{4})'),
        'Communication Address': re.compile(r'Communication Address & Place of Supply\s*:\s*([\s\S]*?)\s*Policy Issuing Branch'),
        'Policy Issuing Branch': re.compile(r'Policy Issuing Branch\s*:\s*([\s\S]*?)\s*Mobile No'),
        'Mobile No': re.compile(r'Mobile No\s*:\s*(\d{4}\*\*\*\*\*\*)'),
        'Tax Invoice No': re.compile(r'Tax Invoice No. & Date\s*:\s*(\S+)\s*&\s*(.*?)(?=\n|$)'),
        'Email-ID': re.compile(r'Email-ID\s*:\s*(\S+@\S+\.\S+)'),
        'Registration No.': re.compile(r'Registration\s+No\.\s+([A-Z]{2}\d{2}[A-Z]{2}\d{4})'),
        'GSTIN/UIN & Place of Supply': re.compile(r'GSTIN/UIN & Place of Supply\s*:\s*(.*?)(?=\n|$)'),
        'Nominee Name': re.compile(r'Nominee Name\s*:\s*(.*?)\s*$'),
        'Mfg. Month & Year': re.compile(r'Mfg. Month & Year\s*\s*(\w+-\d{4})'),
        'Make / Model & Variant': re.compile(r'Make / Model & Variant\s*:\s*(.*?)\s*(?=CC / HP / Watt|$)'),
        'CC / HP / Watt': re.compile(r'CC / HP / Watt\s*\s*(\d+)'),
        'Engine No.': re.compile(r"Engine\s+No\.\s*/\s*Chassis\s+No\.\s*(\S+)\s*/\s*(\S+)"),
        'Seating Capacity of side car': re.compile(r'Seating Capacity of side car \(if any\)\s+Including driver\s+(\d+)'),


        'Type of Body / LCC': re.compile(r'Type\s*of\s*Body\s*/\s*LCC\s*([A-Z]+\s*/\s*\d+)$'),


        'Total Premium (₹)': re.compile(r'Total Premium\s*\(\s*₹\s*\)\s*(\d+(?:\.\d{1,2})?)'),
        'RTO Location': re.compile(r'RTO Location\s*\s*(.*?)(?=\n|$)'),
        'Total IDV (₹)': re.compile(r'Total IDV\s*\(\s*₹\s*\)\s*([\d,]+\.\d{2})'),
        'Hypothecation/Lease': re.compile(r'Hypothecation/Lease\s*\s*(.*?)(?=\n|$)'),

        'Total Basic Liability Premium': re.compile(r'Total Basic Liability Premium\s*([\d,]+\.\d{2})'),
        'TOTAL LIABILITY PREMIUM': re.compile(r'TOTAL LIABILITY PREMIUM\s*([\d,]+\.\d{2})'),

        'TOTAL OWN DAMAGE PREMIUM': re.compile(r'TOTAL OWN DAMAGE PREMIUM\s*(\d+(?:\.\d{1,2})?)'),
        # 'Liability - Section II Amount': re.compile(r'Liability - Section II Amount\s*\s*(\d+\.\d{2})'),
        # 'PA Benefits - Section III': re.compile(r'TOTAL LIABILITY PREMIUM\s*\s*(\d+\.\d{2})'),
        'TOTAL PACKAGE PREMIUM': re.compile(r'TOTAL PACKAGE PREMIUM\s*\(\s*Sec\s*I\s*\+\s*II\s*\+\s*III\s*\)\s*([\d,]+\.\d{2})'),
        'CGST': re.compile(r'CGST \(@9.00%\)\s*(\d+\.\d{2})'),
        'SGST': re.compile(r'SGST \(@9\.00%\)\s*(\d+\.\d{2})'),
        'IGST (@18.00%)': re.compile(r'IGST \(@18\.00%\)\s*([\d,]+\.\d{2})'),
        'TOTAL PREMIUM PAYABLE': re.compile(r'TOTAL PREMIUM PAYABLE\s*\(\s*₹\s*\)\s*([\d,]+\.\d{2})'),
    }

    extracted_data = {}
    for key, pattern in patterns.items():
        match = pattern.search(text)
        if match:
            extracted_data[key] = match.group(1) if match.lastindex else match.group(0)
        else:
            extracted_data[key] = None

    return extracted_data


pdf_file = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\RELIANCE PDF\110722423470031376.pdf'

# Extract data from the PDF
try:
    data = extract_data_from_pdf(pdf_file)
    # Print the extracted data
    for key, value in data.items():
        print(f"{key}: {value}")
except Exception as e:
    print(f"Error processing {pdf_file}: {e}")