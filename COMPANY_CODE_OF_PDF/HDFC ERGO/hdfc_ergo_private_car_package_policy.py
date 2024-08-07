import re
import fitz

patterns = {
    'Company_name_pattern': re.compile(r'\bHDFC ERGO General Insurance Company Limited\b'),
    "Policy Type": re.compile(r'Certificate of Insurance cum Policy Schedule\s+([^\n]+)', re.MULTILINE),
    'Name': re.compile(r'Email ID\s*:\s*[\w\.-]+@[\w\.-]+\s*\n(.*)', re.MULTILINE),
    'Policy Number': re.compile(r'Policy\s+No\.\s+(\d{4}\s\d{4}\s\d{4}\s\d{4}\s\d{3})'),
    'Intermediary Name': re.compile(r'BROKER Name\s*:\s*(.*?)\s*BROKER Code'),
    'Policy Start Date': re.compile(r'Period of\s+Insurance\s+From\s+(\d{2} [A-Za-z]+, \d{4})'),
    'Policy End Date': re.compile(r'To\s+(\d{2} [A-Za-z]+, \d{4})'),
    'Receipt Date': re.compile(r'Issuance\s+Date\s+(\d{2}/\d{2}/\d{4})'),
    'Mobile No': re.compile(r'Tel\.\s+(\d+)'),
    'Email': re.compile(r'Email ID\s+:\s+([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})'),
    # 'Address': re.compile(r'MR\s+[A-Z\s]+(?:\n[A-Z\s\d-]+)+\n[A-Z\s-]+\s+\d{6}\nTel\.\s+\d+', re.MULTILINE),
    'Registration Number': re.compile(r'Registration No\s+([A-Z0-9-]+)'),
    'RTO': re.compile(r'RTO\s+([\w]+)'),
    'Engine Number': re.compile(r'Engine No\.\s+(\S+)'),
    'Chassis Number': re.compile(r'Chassis No\.\s+(\S+)'),
    'Year of Manufacture': re.compile(r'Year of Manufacture\s+(\d{4})'),
    'Make': re.compile(r'Make\s+([A-Z]+)', re.IGNORECASE),
    'Model': re.compile(r'Model\s+-\s*([A-Za-z0-9\s\(\)\.]+)'),
    'Body Type': re.compile(r'Body Type\s+([A-Za-z]+)'),
    'CC': re.compile(r'Cubic Capacity\s*/\s*Watts\s*(\d+)'),
    'Seating Capacity including Driver': re.compile(r'Seats\s*(\d+)'),
    'Previous policy number': re.compile(r'Previous Policy No\.\s*([\w\-\/]+)'),
    'Previous policy end date': re.compile(r'Valid From\s*(\d{2}/\d{2}/\d{4})\s+to\s+(\d{2}/\d{2}/\d{4})'),
    'Geographical Area': re.compile(r'Geographical Area\s+(\w+)'),
    'Total IDV': re.compile(r'Total IDV\s+([\d,]+)'),
    'Total Basic Premium': re.compile(r'Basic Own Damage:\s*(\d+)'),
    'Net Own Damage Premium(a)': re.compile(r'Net Own Damage Premium \(a\)\s+(\d+)'),
    'Net Liability Premium (b)': re.compile(r'Net Liability Premium \(b\)\s+(\d+)'),
    'Total Package Premium (a+b)': re.compile(r"Total Package Premium \(a\+b\)\s*(\d+)"),
    'GST': re.compile(r'Integrated Tax 18%\s+(\d+)'),
    'Total Premium': re.compile(r'Total Premium\s+([\d,.]+)'),

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
            try:
                details[key] = match.group(1)
            except IndexError:
                details[key] = None
        else:
            details[key] = None
    return details


# Path to the uploaded PDF file
pdf_path = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\HDFC ERGO\2302101448409600000.pdf'
# Extract text from the PDF
text = extract_text_from_pdf(pdf_path)

# Extract details using the updated patterns
details = extract_details(text, patterns)

# Print the extracted details
for key, value in details.items():
    print(f"{key}: {value}")




