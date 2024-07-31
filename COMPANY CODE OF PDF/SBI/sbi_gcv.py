import re
import fitz
import pandas as pd
import os

# Define and compile the regex patterns with capturing groups
patterns = {
    "Policy Number": re.compile(r"Policy Number\s*:\s*([A-Z0-9]+)", re.IGNORECASE),
    "Name": re.compile(r"Insured Name\s*:\s*(.*)", re.IGNORECASE),
    "Address": re.compile(r"Address\s*:\s*(.*)", re.IGNORECASE),
    "Mobile_number": re.compile(r"(\+91-\d{10})", re.IGNORECASE),
    "Make_Model": re.compile(r"Make& Model\s*(.*)", re.IGNORECASE),
    "Year_of_manufacturing": re.compile(r"Year of Manufacturing\s*(\d{4})", re.IGNORECASE),
    'Registration Number':re.compile(r'Registration\s+Number\s+([A-Z0-9]+)'),
    "Engine Number": re.compile(r'Engine\s+Number\s+(\S+)', re.IGNORECASE),
    "Chassis Number": re.compile(r'Chassis Number\s*(.*)', re.IGNORECASE),
    "gross_vehicle_weight": re.compile(r"Gross Vehicle Weight \(GVW\)\s*(\d+)", re.IGNORECASE),
    "Seating_capacity": re.compile(r"Carrying Capacity\s*(\d+)", re.IGNORECASE),
    "Vehicle_body_type": re.compile(r"Vehicle Body Type\s*[:\s]*([^\r\n]+)", re.IGNORECASE),
    "RTO": re.compile(r"RTO Location Name\s*(.*)", re.IGNORECASE),
    'Intermediary Name': re.compile(r"Intermediary Name\s*:\s*(.*)", re.IGNORECASE),
    'Basic TP Premium': re.compile(r"Basic TP Premium\s*([\d,]+\.\d{2})", re.IGNORECASE),
    "total_liability_premium": re.compile(r"B\)\s*TOTAL LIABILITY PREMIUM\s*([\d,]+\.\d{2})", re.IGNORECASE),
    "total_policy_premium": re.compile(r"Total Policy Premium \(A\+B\)\s*([\d,]+\.\d{2})", re.IGNORECASE),
    "taxes_as_applicable": re.compile(r"Taxes as applicable\s*([\d,]+\.\d{2})", re.IGNORECASE),
    "kerala_flood_cess": re.compile(r"Kerala Flood Cess @ 1%\s*([\d,]+\.\d{2})", re.IGNORECASE),
    "total_premium_collected": re.compile(r"Total Premium Collected\s*([\d,]+\.\d{2})", re.IGNORECASE),
    "policy_start_date": re.compile(r"Policy Start Date\s*(\d{4}-\d{2}-\d{2})"),
    "policy_end_date": re.compile(r"Policy End Date\s*(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})"),
    "type_of_cover": re.compile(r"Type of Cover\s*(.*)", re.IGNORECASE),
}

def extract_text_from_pdf(pdf_path):
    """
    Extracts text from a PDF file.
    Args:
        pdf_path (str): Path to the PDF file.
    Returns:
        str: Extracted text or an empty string if an error occurs.
    """
    try:
        document = fitz.open(pdf_path)
        text = ""
        for page_num in range(len(document)):
            page = document.load_page(page_num)
            text += page.get_text()
        return text
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")
        return ""

def extract_details(text, patterns):
    """
    Extracts details from the provided text using predefined regex patterns.
    Args:
        text (str): Text from the PDF.
        patterns (dict): Dictionary of regex patterns.
    Returns:
        dict: Extracted details for each pattern key.
    """
    details = {}
    if text:
        for key, pattern in patterns.items():
            try:
                match = pattern.search(text)
                details[key] = match.group(1).strip() if match and match.lastindex else None
            except Exception as e:
                print(f"Error extracting '{key}': {e}")
                details[key] = None
    else:
        print("No text available for extraction.")

    return details

def process_pdfs_in_folder(folder_path):
    """
    Processes all PDF files in a given folder and extracts data based on patterns.
    Args:
        folder_path (str): Path to the folder containing PDF files.
    Returns:
        DataFrame: DataFrame containing extracted details from all PDFs.
    """
    all_details = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing file: {pdf_path}")
            text = extract_text_from_pdf(pdf_path)
            details = extract_details(text, patterns)
            details['Filename'] = filename  # Add filename to the details
            all_details.append(details)

    # Create a DataFrame from the list of details
    df = pd.DataFrame(all_details)
    return df

def save_to_excel(df, output_path):
    """
    Saves the DataFrame to an Excel file.
    Args:
        df (DataFrame): DataFrame to be saved.
        output_path (str): Path to the output Excel file.
    """
    try:
        df.to_excel(output_path, index=False)
        print(f"Data successfully saved to {output_path}")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

# Folder path containing PDF files
folder_path = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\SBI PDF'
excel_path = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\SBI PDF\sbi_gcv.xlsx'

# Process PDFs and get the DataFrame
df = process_pdfs_in_folder(folder_path)

# Save the DataFrame to an Excel file
save_to_excel(df, excel_path)
