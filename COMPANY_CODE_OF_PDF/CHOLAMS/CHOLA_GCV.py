import re
import pdfplumber

# Path to the PDF file
pdf_file_path = r"C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\CHOLAMS\33790394772500000.pdf"

# Open the PDF file using pdfplumber
text = ''
try:
    with pdfplumber.open(pdf_file_path) as pdf:
        for page in pdf.pages:
            # Extract text from each page and concatenate
            text += page.extract_text() or ''
except Exception as e:
    print(f"Error opening or reading the PDF file: {e}")
    text = ''

# Regular expressions to extract data
patterns = {
    'policy_type': r'Motor\s+([^U]+?)\s+UIN',
    'policy_number': r'Policy cum Certificate Number\s*([0-9A-Z/]+)',
    'insured_name': r'Name & Communication Address of the Insured:\s*([A-Za-z\s]+)',
    'customer_phone_number': r'Mobile/Landline No\s*:\s*([\d\s-]+)',
    'insured_address': r'Registration Address:\s*([A-Za-z\s\S]*?\d{6})',
    'date_of_issuance': r'Receipt Date:\s*([\d/-]+)',
    'period_of_insurance_from': r'Period of Insurance: From\s*([\d/-]+)',
    'period_of_insurance_to': r'To: Midnight of\s*([\d/-]+)',
    'registration_number': r'Registration Mark\s*([\w\d]+)',
    'date_of_registration': r'Date of Registration\s*([\d/-]+)',
    'make': r'Make\s+([A-Z0-9\s]+)(?=\s*Model|$)',
    'model': r'Model\s+([A-Z0-9\.\s-]+)(?=\s*Variant|$)',
    'variant': r'Variant\s+([A-Z0-9\s\.-]+)(?=\s*Year|$)',
    'engine_number': r'Engine Number\s*([\w\d]+)',
    'chassis_number': r'Chassis Number\s*([\w\d]+)',
    'seating_capacity': r'Total seating capacity including driver\s*([\d]+)',
    'cc': r'Cubic Capacity\s*/\s*KW\s*([\d]+)',
    'previous_insurer_name': r'Previous Insurer\s*([\w\s]+)',
    'previous_policy_no': r'Previous Policy No\s*([\w\d/]+)',
    'previous_policy_expiry_date': r'Previous Policy Expiry Date\s*([\d/-]+)',
    'total_idv': r'For Vehicle\s*([\d]+)',
    'total_od_premium': r'Basic OD\s*[\d]+\s*([\d]+)',
    'third_party_liability': r'Basic TP\s*([\d]+)',
    'net_premium': r'TOTAL PREMIUM \(A\)\s*([\d]+)',
    'gst': r"IGST\s*\(\d{1,2}%\)\s*(\d+)",
    'final_premium':r'TOTAL AMOUNT Rs\.\s*(\d+[\d,]*)'
}

# Extract the data using regex
extracted_data = {}
for key, pattern in patterns.items():
    match = re.search(pattern, text)
    if match:
        extracted_data[key] = match.group(1).strip()
    else:
        extracted_data[key] = 'Not found'

# Print the extracted data
for key, value in extracted_data.items():
    print(f"{key}: {value}")
