import pandas as pd
import fitz  # PyMuPDF
import re
import os

def extract_text_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text


def extract_name(text):
    return extract_field_value(text, r"Insured's Name:\s*(.+?)\s*Insured's GST No.")



def extract_father_name(text):
    father_pattern = re.compile(r"S/O\s+(\w+)", re.IGNORECASE)
    match = father_pattern.search(text)
    if match:
        return match.group(1) if match.group(1) else match.group(2)
    return None



def extract_date_of_issue(text):
    date_pattern = re.compile(r"Date of Issue:\s*(\d{2}/\d{2}/\d{4})")
    match = date_pattern.search(text)
    if match:
        return match.group(1)
    return None

def extract_phone_number(text):
    phone_pattern = re.compile(r'\+?\d{10,12}')  # Adjust pattern to match phone numbers
    match = phone_pattern.search(text)
    if match:
        return match.group(0)
    return None

patterns = {}



def extract_period_of_insurance(text):
    period_pattern = re.compile(r"Period of Insurance:\s*(\d{2}/\d{2}/\d{4})\s*to\s*(\d{2}/\d{2}/\d{4})")
    match = period_pattern.search(text)
    if match:
        return match.group(1), match.group(2)
    return None, None



def extract_policy_number(text):
    policy_pattern = re.compile(r"Policy\s+No\.\s*(\d{5,30})", re.IGNORECASE)
    match = policy_pattern.search(text)
    if match:
        return match.group(1)
    return None


def extract_pan_number_from_pdf(pdf_path):
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_path)

        # Regular expression pattern for PAN number
        pan_pattern = re.compile(r'\b[A-Z]{5}[0-9]{4}[A-Z]{1}\b', re.IGNORECASE)

        # Iterate through each page and extract text
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()

            # Search for PAN number in the text
            match = pan_pattern.search(text)
            if match:
                return match.group(0)  # Return the matched PAN number

        # If PAN number is not found
        return None

    except Exception as e:
        print(f"Error: {e}")
        return None

    finally:
        if pdf_document:
            pdf_document.close()


#for the GST number
def extract_gst_number_from_pdf(pdf_path):
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_path)

        # Regular expression pattern for GST number
        gst_pattern = re.compile(r'\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}\d{1}[A-Z]{1}\d{1}\b', re.IGNORECASE)

        # Iterate through each page and extract text
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()

            # Search for GST number in the text
            match = gst_pattern.search(text)
            if match:
                return match.group(0)  # Return the matched GST number

        # If GST number is not found
        return None

    except Exception as e:
        print(f"Error: {e}")
        return None

    finally:
        if pdf_document:
            pdf_document.close()


# For the Insured's ID:
def extract_insured_id(text):
    id_pattern = re.compile(r"Insured's ID: (\w+)", re.IGNORECASE)
    match = id_pattern.search(text)
    if match:
        return match.group(1)
    return None

#For the  Policy Servicing Office
def extract_policy_servicing_office(text):
    office_pattern = re.compile(r"Policy Servicing Office: (\w+)", re.IGNORECASE)
    match = office_pattern.search(text)
    if match:
        return match.group(1)
    return None

# For zone and geographical area
def extract_zone_and_geographical_area(text):
    zone_pattern = re.compile(r"Zone: (\w+)", re.IGNORECASE)
    area_pattern = re.compile(r"Geographical Area: (\w+)", re.IGNORECASE)
    zone_match = zone_pattern.search(text)
    area_match = area_pattern.search(text)
    if zone_match and area_match:
        return zone_match.group(1), area_match.group(1)
    return None, None


def extract_period_of_insurance(text):
   period_pattern = re.compile(r"Period of Insurance:\s*From\s*(\d{2}:\d{2}:\d{2})?\s*(?:of)?\s*(\d{2}/\d{2}/\d{4})\s*to\s*(\d{2}:\d{2}:\d{2})?\s*(?:of)?\s*(\d{2}/\d{2}/\d{4})")
   match = period_pattern.search(text)
   if match:
        start_time = match.group(1) or "00:00:00"
        start_date = match.group(2)
        end_time = match.group(3) or "23:59:59"
        end_date = match.group(4)
        return f"{start_date} {start_time}", f"{end_date} {end_time}"
   return None, None



def extract_field_value(text, field_pattern):
    match = re.search(field_pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


# For the make model and varient
def extract_vehicle_details(pdf_path):
    # Open the PDF file
    document = fitz.open(pdf_path)


    # Initialize variables to hold the extracted details
    make = model = variant = cc = year_of_manufacture = seating_capacity = engine_number = chassis_number = None

    # Iterate through each page
    try:
        # Open the PDF file
        document = fitz.open(pdf_path)

        # Iterate through each page
        for page_num in range(document.page_count):
            page = document.load_page(page_num)
            text = page.get_text()

            # Split the text into lines for processing
            lines = text.split('\n')

            # Iterate through the lines to find the vehicle details
            for i, line in enumerate(lines):
                if "Make" in line:
                    make = extract_detail(lines, i)
                elif "Model" in line:
                    model = extract_detail(lines, i)
                    variant = extract_detail(lines, i + 1)  # Look for variant in the next line
                    if '&' in model:
                        model, variant = map(str.strip, model.split('&', 1))

                    else:
                        model = variant
                elif "Year of Manufacture" in line:
                    year_of_manufacture = extract_detail(lines, i)
                elif "Seating Capacity" in line:
                    seating_capacity = extract_detail(lines, i)
                elif "Engine No." in line or "Engine Number" in line:
                    engine_number = extract_detail(lines, i)
                elif "Chassis No." in line or "Chassis Number" in line:
                    chassis_number = extract_detail(lines, i)

        # Close the document
        document.close()

    except Exception as e:
        print(f"Error extracting details from {pdf_path}: {e}")

    return make, model, variant, cc, year_of_manufacture, seating_capacity, engine_number, chassis_number

#For the cc
def extract_combined_info(pdf_path):
    # Open the PDF file
    document = fitz.open(pdf_path)

    # Initialize variable to hold the extracted detail
    combined_info = None

    # Iterate through the pages
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text = page.get_text()

        # Search for the combined term
        combined_index = text.find('Cubic Capacity / GVW / HP')
        if combined_index != -1:
            combined_line = text[combined_index:combined_index+40]
            combined_info = ''.join(filter(lambda x: x.isdigit() or x.isspace(), combined_line))

    # Return the extracted detail
    return combined_info



def extract_detail(lines, index):
    # Extract detail by handling potential variations
    detail = None
    if index < len(lines) - 1:
        detail = lines[index + 1].strip()
    return detail

def parse_text(text):
    data = {}
    # Define regex patterns for each field
    patterns = {
        "Fuel Type": r"Fuel Type\s+([A-Za-z]+)",
        "Body type": r"Body type\s+([A-Za-z]+)",
        "Year of Manufacture": r"Year of Mahufacture\s+(\d{4})",
        "Transmission Type": r"Transmission Type\s+([A-Za-z]+)"
    }

    # Extract data using regex
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            data[key] = match.group(1)
    return data

# Function to extract a specific section from the PDF
def extract_section(pdf_path):
    # Open the PDF file
    document = fitz.open(pdf_path)
    output = ""

    # Initialize variables to store extracted details
    product_name = product_type = business_type = ""


    # Iterate through each page
    for page_num in range(document.page_count):
        page = document.load_page(page_num)
        text = page.get_text("text")

        # Search for each detail using regular expressions
        product_name_match = re.search(r'Product Name\s+(.*)', text, re.IGNORECASE)
        product_type_match = re.search(r'Product Type\s+(.*)', text, re.IGNORECASE)
        business_type_match = re.search(r'Business Type\s+(.*)', text, re.IGNORECASE)

        # Extract details if matches are found

        if product_name_match:
            product_name = product_name_match.group(1).strip()
            product_name = clean_text(product_name)
        if product_type_match:
            product_type = product_type_match.group(1).strip()
        if business_type_match:
            business_type = business_type_match.group(1).strip()

        # If all details are found, break out of the loop
        if  product_name and product_type and business_type:
            break

    # Construct the formatted output

    output += f"Product Name: {product_name}\n, "
    output += f"Product Type: {product_type}\n, "
    output += f"Business Type: {business_type}\n"

    document.close()
    return output
def clean_text(text):
    return text.replace("Valid to (dd-mmm-yyyy)", "").strip()


#For the liability
def extract_insurance_data(pdf_path):
    data = {}
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()

    # Extract base premium
    base_premium_match = re.search(r"Base Premium including Premium for TPPD\s+₹\s*([0-9,.]+)", text)
    if base_premium_match:
        data["base_premium"] = base_premium_match.group(1)

    # Extract legal liability to paid drivers
    legal_liability_match = re.search(r"Legal Liability To Paid Drivers\s+₹\s*([0-9,.]+)", text)
    if legal_liability_match:
        data["legal_liability_to_paid_drivers"] = legal_liability_match.group(1)

    # Extract total liability premium
    total_liability_premium_match = re.search(r"Total liability premium\s+₹\s*([0-9,.]+)", text)
    if total_liability_premium_match:
        data["total_liability_premium"] = total_liability_premium_match.group(1)

    # Extract CGST
    cgst_match = re.search(r"CGST @\d+%\s+₹\s*([0-9,.]+)", text)
    if cgst_match:
        data["cgst"] = cgst_match.group(1)

    # Extract SGST
    sgst_match = re.search(r"SGST @\d+%\s+₹\s*([0-9,.]+)", text)
    if sgst_match:
        data["sgst"] = sgst_match.group(1)

    # Extract final premium
    final_premium_match = re.search(r"Final premium\s+₹\s*([0-9,.]+)", text)
    if final_premium_match:
        data["final_premium"] = final_premium_match.group(1)

    return data

# For the policy payment


def extract_product_name(text):
    return extract_field_value(text, r"Product Name:\s*(.+)")


#Define previous details
previous_insurer_pattern = re.compile(r'Previous Insurer\s*:\s*(.*?)\s*\n', re.DOTALL | re.IGNORECASE)

# Define the regex pattern for Previous Policy No.
previous_policy_no_pattern = re.compile(r'Previous Policy No\.\s*:\s*(.*?)\s*\n', re.DOTALL | re.IGNORECASE)

# Define the regex pattern for Expiry on
expiry_on_pattern = re.compile(r'Expiry on\s*:\s*(.*?)\s*\n', re.DOTALL | re.IGNORECASE)

def extract_relationship_of_appointee_with_nominee(pdf_path):
    try:
        pdf_document = fitz.open(pdf_path)
        relationship_pattern = re.compile(r"Relationship of Appointee with Nominee:\s*(\w+)", re.IGNORECASE)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()
            match = relationship_pattern.search(text)
            if match:
                return match.group(1)
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None
    finally:
        if pdf_document:
            pdf_document.close()



# Function to extract data from a single PDF
def extract_data_from_pdf(pdf_path):
    try:
        # Extract data using your existing functions
        pdf_text = extract_text_from_pdf(pdf_path)

        name = extract_name(pdf_text)
        father_name = extract_father_name(pdf_text)
        phone_number = extract_phone_number(pdf_text)
        date_of_issue = extract_date_of_issue(pdf_text)
        policy_number = extract_policy_number(pdf_text)
        start_date, end_date = extract_period_of_insurance(pdf_text)
        address = extract_field_value(pdf_text, r"Address: (.+?)\n")
        pan_number = extract_pan_number_from_pdf(pdf_path)
        gst_number = extract_gst_number_from_pdf(pdf_path)
        insured_id = extract_insured_id(pdf_text)
        geographical_area, policy_servicing_office = extract_zone_and_geographical_area(pdf_text)

        relationship_appointee_nominee = extract_relationship_of_appointee_with_nominee(pdf_path)
        make, model, variant, cc, year_of_manufacture, seating_capacity, engine_number, chassis_number = extract_vehicle_details(pdf_path)
        combined_info = extract_combined_info(pdf_path)
        extracted_data = extract_insurance_data(pdf_path)
        parsed_data = parse_text(pdf_text)
        extracted_details = extract_section(pdf_path)
        product_name = extract_product_name(pdf_path)

        # For previous insurance details
        previous_insurer_match = previous_insurer_pattern.search(pdf_text)
        previous_insurer = previous_insurer_match.group(1).strip() if previous_insurer_match else None

        previous_policy_no_match = previous_policy_no_pattern.search(pdf_text)
        previous_policy_no = previous_policy_no_match.group(1).strip() if previous_policy_no_match else None

        expiry_on_match = expiry_on_pattern.search(pdf_text)
        expiry_on = expiry_on_match.group(1).strip() if expiry_on_match else None

        # For the liability
        data = extract_insurance_data(pdf_path)


        # Prepare data dictionary
        data_dict = {
            "Name": name,
            "Father's Name": father_name,
            "Contact Number": phone_number,
            "Date of Issue": date_of_issue,
            "Policy Number": policy_number,
            "Period of Insurance Start": start_date,
            "Period of Insurance End": end_date,
            "Address": address,
            "PAN Number": pan_number,
            "GST Number": gst_number,
            "Insured ID": insured_id,
            "Geographical Area": geographical_area,
            "Policy Servicing Office": policy_servicing_office,
            "Relationship of Appointee with Nominee": relationship_appointee_nominee,
            "Make": make,
            "Model": model,
            "Variant": variant,
            "Cubic Capacity / GVW / HP": combined_info,
            "Year of Manufacture": year_of_manufacture,
            "Seating Capacity": seating_capacity,
            "Engine Number": engine_number,
            "Chassis Number": chassis_number,
            #"Product Name": product_name,
            # Additional fields
            "Fuel Type": parsed_data.get('Fuel Type', ''),
            "Body type": parsed_data.get('Body type', ''),
            "Transmission Type": parsed_data.get('Transmission Type', ''),
            "Base Premium": extracted_data.get('base_premium', ''),
            "Legal Liability To Paid Drivers": extracted_data.get('legal_liability_to_paid_drivers', ''),
            "Total Liability Premium": extracted_data.get('total_liability_premium', ''),
            "CGST": extracted_data.get('cgst', ''),
            "SGST": extracted_data.get('sgst', ''),
            "Final Premium": extracted_data.get('final_premium', ''),
            "Previous Insurer": previous_insurer,
            "Previous Policy No.": previous_policy_no,
            "Expiry on": expiry_on,
            "Section": extracted_details,



            # Add more fields as needed


        }

        return data_dict

    except Exception as e:
        print(f"Error extracting data from {pdf_path}: {e}")
        return None


# List of PDF files to process
pdf_files = [

    r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\ZUNO PDF\402000209545.pdf',
    r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\ZUNO PDF\402000209736.pdf',
    r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\ZUNO PDF\402000212205.pdf',
    r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\ZUNO PDF\402000221020.pdf',
    r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\ZUNO PDF\402000221602.pdf',






    # Add more PDF paths as needed
]

# List to store extracted data from all PDFs
all_data = []

# Iterate through each PDF file
for pdf_file in pdf_files:
    # Extract data from each PDF
    extracted_data = extract_data_from_pdf(pdf_file)

    if extracted_data:
        all_data.append(extracted_data)

# Convert list of dictionaries to a pandas DataFrame
df = pd.DataFrame(all_data)

# Define the output Excel file path
output_excel = r'C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\ZUNO PDF\ZUNO_EXCEL.xlsx'

# Save the DataFrame to Excel
df.to_excel(output_excel, index=False)

print(f"Extracted ZUNO data saved to {output_excel}")

