import fitz  # PyMuPDF
import re
import os
import pandas as pd


# Function to extract data from a single PDF
def extract_data_from_pdf(pdf_path):
    extracted_data = {}

    try:
        # Try opening the PDF file
        document = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening the PDF file '{pdf_path}': {e}")
        return None

    text = ""

    try:
        # Try reading the text from each page
        for page_num in range(len(document)):
            try:
                page = document.load_page(page_num)
                text += page.get_text()
            except Exception as e:
                print(f"Error reading page {page_num} from '{pdf_path}': {e}")
                continue
    except Exception as e:
        print(f"Error processing the document '{pdf_path}': {e}")
        return None

    # Define regex patterns
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
        'Registration No.': re.compile(r'Registration\s+No\.\s+([A-Z]{2}\d{2}[A-Z]{2}\d{4})'),
        'GSTIN/UIN & Place of Supply': re.compile(r'GSTIN/UIN & Place of Supply\s*:\s*(.*?)(?=\n|$)'),
        'Mfg. Month & Year': re.compile(r'Mfg. Month & Year\s*\s*(\w+-\d{4})'),
        'Make / Model & Variant': re.compile(r'Make\s*/\s*Model\s*&\s*Variant\s*(.+)'),
        'CC / HP / Watt': re.compile(r'CC / HP / Watt\s*\s*(\d+)'),
        'Engine No.': re.compile(r"Engine\s+No\.\s*/\s*Chassis\s+No\.\s*(\S+)\s*/\s*(\S+)"),
        'Chassis No.': re.compile(r'Chassis\s+Number\s+(\S+)'),
        'Seating Capacity': re.compile(r'LCC\(excluding\s+driver\)\s*(\d+)'),
        'Type of Body': re.compile(r'Type\s+of\s+Body\s*([^\n]*)'),
        'Total Premium': re.compile(r'Total\s+Premium\s*\(\s*`\s*\)\s*([\d,]+(?:\.\d{2})?)'),
        'RTO Location': re.compile(r'RTO Location\s*\s*(.*?)(?=\n|$)'),
        'Total IDV': re.compile(r'Total\s+IDV\s*\(`\)\s*(\w+)'),
        'Hypothecation/Lease': re.compile(r'Hypothecation/Lease\s*\s*(.*?)(?=\n|$)'),
        'TOTAL OWN DAMAGE PREMIUM': re.compile(r'TOTAL OWN DAMAGE PREMIUM\s*(\d+(?:\.\d{1,2})?)'),
        'Total Basic Liability Premium': re.compile(r'Total\s+Basic\s+Liability\s+Premium\s*([\d,]+\.\d{2})'),
        'TOTAL LIABILITY PREMIUM': re.compile(r'TOTAL\s+LIABILITY\s+PREMIUM\s*([\d,]+\.\d{2})'),
        'TOTAL PACKAGE PREMIUM': re.compile(
            r'TOTAL\s+PACKAGE\s+PREMIUM\s*\(\s*Sec\s+I\s*\+\s*II\s*\+\s*III\s*\)\s*([\d,]+\.\d{2})'),
        'CGST': re.compile(r'CGST\s*\(@9\.00%\)\s*([\d,]+\.\d{2})'),
        'SGST': re.compile(r'SGST\s*\(@9\.00%\)\s*([\d,]+\.\d{2})'),
        'IGST (@18.00%)': re.compile(r'IGST\s*\(@18\.00%\)\s*([\d,]+\.\d{2})'),
        'TOTAL PREMIUM PAYABLE': re.compile(r'TOTAL PREMIUM PAYABLE\s*\(`\)\s*([\d,]+\.\d{2})')
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


# Function to process all PDFs in a folder and save results to Excel
def process_pdfs_in_folder(folder_path, output_excel_file):
    all_data = []

    # Try processing each PDF file in the folder
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            try:
                data = extract_data_from_pdf(pdf_path)
                if data:
                    data['Filename'] = filename
                    all_data.append(data)
            except Exception as e:
                print(f"Error processing {filename}: {e}")

    # Save the extracted data to an Excel file
    try:
        df = pd.DataFrame(all_data)
        df.to_excel(output_excel_file, index=False)
        print(f"Data successfully saved to {output_excel_file}")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")


# Example usage
pdf_folder = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\RELIANCE PDF'
output_excel = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\RELIANCE PDF\reliance_miss-D.xlsx'

try:
    process_pdfs_in_folder(pdf_folder, output_excel)
except Exception as e:
    print(f"Error processing PDF folder: {e}")
