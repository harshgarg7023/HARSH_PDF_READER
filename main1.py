from fastapi import FastAPI, UploadFile, File
from typing import List
import uvicorn
import os
import pandas as pd
from pathlib import Path
import fitz  # PyMuPDF
import re

app = FastAPI()


# Define your extraction function
def extract_text_from_pdf(file_path):
    doc = fitz.open(file_path)
    text = doc[0].get_text()

    patterns = {
        'Policy Number': re.compile(r'Policy Number\s+(\w+)'),
        'Insured Name': re.compile(r'Dear\s+(.*)'),
        'Address': re.compile(r'Registration address of the Insured\s+(.*)'),
        'Mobile No': re.compile(r'Telephone\(Mob\) : (\d+)'),
        'Email': re.compile(r'Email Id : (\S+@\S+\.\S+)'),
        'Intermediary Name': re.compile(r'Intermediary Name\s*:\s*(.*)'),
        'Date of Issue': re.compile(r'Date : (\d{2}/\d{2}/\d{4})'),
        'Policy Start Date': re.compile(r'Risk start time and date\s+(\d{2}/\d{2}/\d{4})'),
        'Policy End Date': re.compile(r'Risk end date\s+(\d{2}/\d{2}/\d{4})'),
        'Place of Supply(State Code)': re.compile(r'Place of Supply\(State Code\): (\d+)'),
        'GSTIN / UIN Number': re.compile(r'GSTIN / UIN Number\s*:\s*([\w-]+)'),
        # 'Address of Service Provider': re.compile(r'Address of Service Provider\s*:\s*(.*?Pincode\s*-\s*\d+)'),
        'Area Code': re.compile(r'Area Code\s*:\s*(.*)'),
        'State Code': re.compile(r'FGI State Code\s*:\s*(\d+)'),
        'FGI GSTIN Number': re.compile(r'FGI GSTIN Number\s*:\s*([\w\d]+)'),
        'FGI PAN Number': re.compile(r'FGI PAN Number:\s*([A-Z0-9]{10})'),
        'Nature of Service': re.compile(r'Nature of Service\s*:\s*(.*)'),
        'Type of Fuel': re.compile(r'Fuel\s+Type\s*(\w+)'),
        'Engine No.': re.compile(r'Engine\s+No\s*\s*(\w+)'),
        'Chassis No.': re.compile(r'Chassis\s+No\s*\s*(\w+)'),
        'Seating Capacity': re.compile(r'Seating Capacity\s*\s*(\d+)'),
        'Cubic Capacity': re.compile(r'Cubic Capacity\s*(\d+)'),
        'Registration Number': re.compile(r'Registration\s+No\s*(\w+)'),
        'Make/Model of Vehicle': re.compile(
            r'Make\s+and\s+Model\s+of\s+vehicle\s+insured\s+([A-Z]+\s+[A-Z]+\s+[A-Z0-9\s-]+)'),
        'Class of Vehicle': re.compile(r'Class\s+of\s+Vehicle\s*:\s*(.*)'),
        'RTO': re.compile(r'RTO\s+where\s+vehicle\s+is/will\s+be\s+registered\s+(.*)'),
        'Year of Manufacture': re.compile(r'Year\s+of\s+Manufacturing\s+(\d{4})'),
        'Nominee Name': re.compile(r'Nominee\s+Name\s+(.*)'),
        'Nominee Relation': re.compile(r'Nominee\s+Relationship\s+with\s+Insured\s+(.*)'),
        'Nominee Age': re.compile(r'Nominee\s+Age\s+in\s+Y\s+or\s+M\s+(\d+Y)'),
        'Nominee %': re.compile(r'Nominee\s+%\s+(\d+)'),
        'Total Own Damage Premium (A)': re.compile(
            r'Total\s+Own\s+Damage\s+Premium\s+\(A\)\s+\(rounded\s+off\)\s+(\d+)'),
        'Basic Premium including Premium for TPPD': re.compile(
            r'Basic\s+Premium\s+including\s+Premium\s+for\s+TPPD\s+([\d,]+\.\d{2})'),
        'Legal Liability to Paid Driver/Cleaner/Employees': re.compile(
            r'Legal\s+Liability\s+to\s+Paid\s+Driver/Cleaner/Employees\s+\(No\. of persons \d+\)\s+([\d,]+\.\d{2})'),
        'Total Liability Premium (B)': re.compile(r'Total\s+Liability\s+Premium\s+\(B\)\s+([\d,]+\.\d{2})'),
        'Total Annual Premium (A+B)': re.compile(r'Total\s+Annual\s+Premium\s+\(A\+B\)\s+([\d,]+\.\d{2})'),
        'Total Premium for the Policy Period': re.compile(
            r'Total\s+Premium\s+for\s+the\s+Policy\s+Period\s+([\d,]+\.\d{2})'),
        'GST': re.compile(r'Goods\s+and\s+Service\s+Tax\s+([\d,]+\.\d{2})'),
        'Final Total Premium ': re.compile(r'Total\s+Premium\s+\(rounded\s+off\)\s+([\d,]+\.\d{2})'),
        'Previous Insurer Name': re.compile(r'Previous Insurer Name\s+(.*)\S+\S+'),
        'Expiring Policy No': re.compile(r'Expiring Policy No\s+(\d+)\S+\S+'),
        'Expiring Policy Expiry Date': re.compile(r'Expiring Policy Expiry Date\s+(\d{2}/\d{2}/\d{4})'),
        'IDV (Insured Declared Value)': re.compile(r'Vehicle\s+IDV\s+on\s+Renewal\s+₹\.(\d+(,\d+)*)')
    }

    extracted_data = {}
    for key, pattern in patterns.items():
        match = pattern.search(text)
        extracted_data[key] = match.group(1) if match else None

    return extracted_data


@app.post("/extract/")
async def extract_data(files: List[UploadFile] = File(...)):
    extracted_data = []
    save_dir = Path(r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\FUTURE PDF')
    save_dir.mkdir(parents=True, exist_ok=True)

    for file in files:
        # Save the uploaded file to the specific directory
        file_location = save_dir / file.filename
        with open(file_location, "wb+") as file_object:
            file_object.write(file.file.read())

        # Call your existing extraction function here
        data = extract_text_from_pdf(file_location)
        extracted_data.append(data)

        # Optionally, remove the file after processing
        os.remove(file_location)

    # Convert extracted data to DataFrame and save to Excel
    df = pd.DataFrame(extracted_data)
    output_file = save_dir / "extracted_data.xlsx"
    df.to_excel(output_file, index=False)

    return {"message": "Data extracted successfully", "output_file": str(output_file)}




