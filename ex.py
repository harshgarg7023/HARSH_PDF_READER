import pdfplumber
import re

# Open the PDF file
pdf_path = "FUTURE COMPANY PDF/VD188386.pdf"
pdf = pdfplumber.open(pdf_path)

# Extract text from all pages
text = ""
for page in pdf.pages:
    text += page.extract_text() + "\n"

# Close the PDF file
pdf.close()

# Define refined regex patterns using re.compile
patterns = {
    "Policy Number": re.compile(r"Policy Number\s*:\s*([A-Z0-9]+)"),
    "Insured Name": re.compile(r"Name of Insured/Proposer\s*:\s*([A-Z\s]+)"),
    "Period of Insurance": re.compile(r"Period of Insurance\s*:\s*From\s*00:00\s*hrs\s*of\s*([0-9/]+)\s*To\s*Midnight\s*of\s*([0-9/]+)"),
    "Engine No.": re.compile(r"Engine No\.\s*:\s*([A-Z0-9]+)"),
    "Chassis No.": re.compile(r"Chassis No\.\s*:\s*([A-Z0-9]+)"),
    "Type of Body": re.compile(r"Type of Body\s*:\s*(\w+)"),
    "Premium": re.compile(r"Premium\s*:\s*₹\s*([\d,]+\.?\d*)"),
    "Cubic Capacity": re.compile(r"Cubic Capacity\s*:\s*(\d+)"),
    "Year of Manufacture": re.compile(r"Year of Manufacture\s*:\s*(\d{4})"),
    "Make/Model of Vehicle": re.compile(r"Make/Model of Vehicle\s*:\s*([\w\s/]+)"),
    "Address": re.compile(r"Address\s*:\s*([^,]+,\s*[^,]+,\s*[^,]+,\s*[A-Z]+\s*,\s*\d{6})"),
    "Registration Number": re.compile(r"Registration Number\s*:\s*([A-Z0-9]+)"),
    "Insurer": re.compile(r"Insurer\s*:\s*([A-Z\s]+)"),
    "Nominee Name": re.compile(r"Nominee Name\s*:\s*([A-Z\s]+)"),
    "Relation with Nominee": re.compile(r"Relation with Nominee\s*:\s*([A-Z\s]+)"),
    "IDV (Insured Declared Value)": re.compile(r"IDV\s*:\s*₹\s*([\d,]+\.?\d*)"),
    "Type of Fuel": re.compile(r"Type of Fuel\s*:\s*(\w+)")


}

# Extract data using regex
extracted_data = {}
for key, pattern in patterns.items():
    match = pattern.search(text)
    if match:
        if key == "Period of Insurance":
            extracted_data[key] = match.group(1)
            extracted_data["Period of Insurance End"] = match.group(2)

        else:
            extracted_data[key] = match.group(1)

# Print extracted data
for key, value in extracted_data.items():
    print(f"{key}: {value}")
