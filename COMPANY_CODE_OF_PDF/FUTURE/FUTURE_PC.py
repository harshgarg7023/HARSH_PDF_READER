import os
import re
import fitz
import pandas as pd


# Define your regex patterns
patterns = {
    'Policy Number': re.compile(r'Policy Number\s+(\w+)'),
    'Insured Name': re.compile(r'Dear\s+(.*)'),
    'Address': re.compile(r'Registration address of the Insured\s+(.*)'),
    'Mobile No': re.compile(r'Telephone\(Mob\) : (\d+)'),
    'Email': re.compile(r'Email Id : (\S+@\S+\.\S+)'),
    'Intermediary Name': re.compile(r'Intermediary Name\s*:\s*(.*)'),
    'Date of Issue': re.compile(r'Date : (\d{2}/\d{2}/\d{4})'),
    'Policy Start Date':re.compile(r'Risk start time and date\s+(\d{2}/\d{2}/\d{4})'),
    'Policy End Date': re.compile(r'Risk end date\s+(\d{2}/\d{2}/\d{4})'),
    'Place of Supply(State Code)': re.compile(r'Place of Supply\(State Code\): (\d+)'),
    'GSTIN / UIN Number': re.compile(r'GSTIN / UIN Number\s*:\s*([\w-]+)'),
    'Area Code': re.compile(r'Area Code\s*:\s*(.*)'),
    'State Code': re.compile(r'FGI State Code\s*:\s*(\d+)'),
    'FGI GSTIN Number': re.compile(r'FGI GSTIN Number\s*:\s*([\w\d]+)'),
    'FGI PAN Number': re.compile(r'FGI PAN Number:\s*([A-Z0-9]{10})'),
    'Nature of Service': re.compile(r'Nature of Service\s*:\s*(.*)'),
    'Type of Fuel': re.compile(r'Fuel\s+Type\s*(\w+)'),
    'Engine No.': re.compile(r'Engine\s+No\s*\s*(\w+)'),
    'Chassis No.': re.compile(r'Chassis\s+No\s*\s*(\w+)'),
    'Seating Capacity': re.compile(r'Seating Capacity\s*\s*(\d+)'),
    'Premium': re.compile(r'Premium\s*\s*([\d,]+\.\d{2})'),
    'Cubic Capacity': re.compile(r'Cubic Capacity\s*(\d+)'),
    'Registration Number': re.compile(r'Registration\s+No\s*(\w+)'),
    'Make/Model of Vehicle': re.compile(r'Make\s+and\s+Model\s+of\s+vehicle\s+insured\s+([A-Z]+\s+[A-Z]+\s+[A-Z0-9\s-]+)'),
    'Class of Vehicle': re.compile(r'Class\s+of\s+Vehicle\s*:\s*(.*)'),
    'RTO': re.compile(r'RTO\s+where\s+vehicle\s+is/will\s+be\s+registered\s+(.*)'),
    'Year of Manufacture': re.compile(r'Year\s+of\s+Manufacturing\s+(\d{4})'),
    'Nominee Name': re.compile(r'Nominee\s+Name\s+(.*)'),
    'Nominee Relation': re.compile(r'Nominee\s+Relationship\s+with\s+Insured\s+(.*)'),
    'Nominee Age': re.compile(r'Nominee\s+Age\s+in\s+Y\s+or\s+M\s+(\d+Y)'),
    'Nominee %': re.compile(r'Nominee\s+%\s+(\d+)'),
    'Basic Premium including Premium for TPPD': re.compile(r'Basic\s+Premium\s+including\s+Premium\s+for\s+TPPD\s+([\d,]+\.\d{2})'),
    'Total Own Damage Premium (A)': re.compile(r'Total\s+Own\s+Damage\s+Premium\s+\(A\)\s+\(rounded\s+off\)\s+([\d,]+\.\d{2})'),
    'Legal Liability to Paid Driver/Cleaner/Employees': re.compile(r'Legal\s+Liability\s+to\s+Paid\s+Driver/Cleaner/Employees\s+\(No\. of persons \d+\)\s+([\d,]+\.\d{2})'),
    'Total Liability Premium (B)': re.compile(r'Total\s+Liability\s+Premium\s+\(B\)\s+([\d,]+\.\d{2})'),
    'Total Annual Premium (A+B)': re.compile(r'Total\s+Annual\s+Premium\s+\(A\+B\)\s+([\d,]+\.\d{2})'),
    'Total Premium for the Policy Period': re.compile(r'Total\s+Premium\s+for\s+the\s+Policy\s+Period\s+([\d,]+\.\d{2})'),
    'GST': re.compile(r'Goods\s+and\s+Service\s+Tax\s+([\d,]+\.\d{2})'),
    'Final Total Premium ': re.compile(r'Total\s+Premium\s+\(rounded\s+off\)\s+([\d,]+\.\d{2})'),
    'Previous Insurer Name': re.compile(r'Previous Insurer Name\s+(.*)\S+\S+'),
    'Expiring Policy No': re.compile(r'Expiring Policy No\s+(\d+)\S+\S+'),
    'Expiring Policy Expiry Date': re.compile(r'Expiring Policy Expiry Date\s+(\d{2}/\d{2}/\d{4})'),
    'IDV (Insured Declared Value)': re.compile(r'Vehicle\s+IDV\s+on\s+Renewal\s+₹\.(\d+(,\d+)*)')

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
            details[key] = match.group(1).strip()
        else:
            details[key] = None
    return details


# Function to process all PDFs in a folder
def process_pdfs_in_folder(folder_path):
    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    all_details = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        text = extract_text_from_pdf(pdf_path)
        details = extract_details(text, patterns)
        details['PDF File'] = pdf_file
        # Add PDF file name to details
        all_details.append(details)

    return all_details



# FUTURE_PC.py

def handle_private_car(data):
    # Process the data for Private Car
    result = f"Processed Private Car data: {data}"
    return result

#FOR THE PROCESS
def process(data):
    # Example processing logic for PC
    # Adjust this based on the actual processing needs
    result = f"Processed PC data: {data['Policy Number']} - {data['Address']}"
    return result



# Path to the folder containing PDFs
folder_path = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\FUTURE PDF'


# Process all PDFs in the folder
extracted_details = process_pdfs_in_folder(folder_path)

# Convert extracted details to DataFrame
df = pd.DataFrame(extracted_details)

# Path to save Excel file
excel_file_path = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\FUTURE PDF\future_private_car.xlsx'

# Save DataFrame to Excel
df.to_excel(excel_file_path, index=False)

print(f"Extracted details from {len(df)} PDFs saved to {excel_file_path}")
