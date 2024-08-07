import fitz  # PyMuPDF
import re
import pandas as pd
import os


# Function to extract data from a single PDF
def extract_data_from_pdf(pdf_path):
    extracted_data = {}

    try:
        # Try opening the PDF file
        document = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening the PDF file: {e}")
        return None

    text = ""

    try:
        # Try reading the text from each page
        for page_num in range(len(document)):
            try:
                page = document.load_page(page_num)
                text += page.get_text()
            except Exception as e:
                print(f"Error reading page {page_num}: {e}")
                continue
    except Exception as e:
        print(f"Error processing the document: {e}")
        return None

    patterns = {
        'Policy Number': re.compile(r'Policy Number\s*:\s*(\S+)'),
        'Covernote No': re.compile(r'Covernote No\s*:\s*(\S+)'),
        'Insured Name': re.compile(r'Insured Name\s*:\s*(.*?)(?=\n|$)'),
        'Period of Insurance Start': re.compile(r'From\s*00:00 Hrs on\s*(\d{2}-\w+-\d{4})'),
        'Period of Insurance End': re.compile(r'to Midnight of\s*(\d{2}-\w+-\d{4})'),
        'Communication Address': re.compile(
            r'Communication Address & Place of Supply\s*:\s*([\s\S]*?)\s*Policy Issuing Branch'),
        'Policy Issuing Branch': re.compile(r'Policy Issuing Branch\s*:\s*([\s\S]*?)\s*Mobile No'),
        'Mobile No': re.compile(r'Mobile No\s*:\s*(\d{4}\*\*\*\*\*\*)'),
        'Tax Invoice No': re.compile(r'Tax Invoice No. & Date\s*:\s*(\S+)\s*&\s*(.*?)(?=\n|$)'),
        'Email-ID': re.compile(r'Email-ID\s*:\s*(\S+@\S+\.\S+)'),
        'Registration No.': re.compile(r'^Registration\s+No\.\s*([A-Z][A-Z\d]*)', re.MULTILINE),
        'GSTIN/UIN & Place of Supply': re.compile(r'GSTIN/UIN & Place of Supply\s*:\s*(.*?)(?=\n|$)'),
        'Mfg. Month & Year': re.compile(r'Mfg. Month & Year\s*\s*(\w+-\d{4})'),
        'Make / Model & Variant': re.compile(r'Make\s*/\s*Model\s*([\w\s]+)\s*/\s*([\w\s]+)\s*/\s*([\w\s]+)'),
        'CC / HP / Watt': re.compile(r'CC / HP / Watt\s*\s*(\d+)'),
        'Engine No.': re.compile(r"Engine\s+No\.\s*/\s*Chassis\s+No\.\s*(\S+)\s*/\s*(\S+)"),
        'Seating Capacity': re.compile(r"LCC Including Driver\s*(\d+)"),
        'Vehicle Category': re.compile(r'Vehicle\s+Category\s*\s*(\w+)'),
        'Type of Body': re.compile(r'Type\s*of\s*Body\s*(\S+)'),
        'RTO Location': re.compile(r'RTO Location\s*\s*(.*?)(?=\n|$)'),
        'Total Premium (₹)': re.compile(r'Total Premium\s*\(\s*₹\s*\)\s*(\d+(?:\.\d{1,2})?)'),
        'Total IDV (₹)': re.compile(r'Total IDV\s*\(\s*₹\s*\)\s*([\d,]+\.\d{2})'),
        'Hypothecation/Lease': re.compile(r'Hypothecation/Lease\s*\s*(.*?)(?=\n|$)'),
        'TOTAL OWN DAMAGE PREMIUM': re.compile(r'TOTAL OWN DAMAGE PREMIUM\s*(\d+(?:\.\d{1,2})?)'),
        'Total Basic Liability Premium': re.compile(r'Total\s+Basic\s+Liability\s+Premium\s*([\d,]+\.\d{2})'),
        'TOTAL LIABILITY PREMIUM': re.compile(r'TOTAL\s+LIABILITY\s+PREMIUM\s*([\d,]+\.\d{2})'),
        'TOTAL PACKAGE PREMIUM': re.compile(
            r'TOTAL PACKAGE PREMIUM\s*\(\s*Sec\s*I\s*\+\s*II\s*\+\s*III\s*\)\s*([\d,]+\.\d{2})'),
        'CGST': re.compile(r'CGST \(@9.00%\)\s*(\d+\.\d{2})'),
        'SGST': re.compile(r'SGST \(@9\.00%\)\s*(\d+\.\d{2})'),
        'IGST (@18.00%)': re.compile(r'IGST \(@18\.00%\)\s*([\d,]+\.\d{2})'),
        'TOTAL PREMIUM PAYABLE': re.compile(r'TOTAL PREMIUM PAYABLE\s*\(\s*₹\s*\)\s*([\d,]+\.\d{2})'),
    }

    for key, pattern in patterns.items():
        try:
            match = pattern.search(text)
            if match:
                extracted_data[key] = match.group(1) if match.lastindex else match.group(0)
            else:
                extracted_data[key] = None
        except Exception as e:
            print(f"Error processing pattern '{key}': {e}")
            extracted_data[key] = None

    return extracted_data


# Function to process all PDFs in a folder and save results to an Excel file
def process_pdfs_in_folder(folder_path, output_file):
    data_list = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing {pdf_path}...")
            try:
                data = extract_data_from_pdf(pdf_path)
                if data:
                    data['File Name'] = filename  # Add the filename to the data
                    data_list.append(data)
            except Exception as e:
                print(f"Error processing {pdf_path}: {e}")

    # Convert the list of data to a DataFrame and save to Excel
    df = pd.DataFrame(data_list)
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Data saved to {output_file}")


# Folder containing PDF files
pdf_folder = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\RELIANCE PDF'

# Output Excel file
output_excel = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\RELIANCE PDF\reliance_pcv.xlsx'

# Process the PDFs and save the results
process_pdfs_in_folder(pdf_folder, output_excel)
