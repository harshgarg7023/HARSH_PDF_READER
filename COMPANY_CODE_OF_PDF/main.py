import fitz  # PyMuPDF
import re
from COMPANY_CODE_OF_PDF.FUTURE import FUTURE_GCV, FUTURE_PC, FUTURE_PCV


def extract_text_from_first_page(pdf_path):
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_path)
        # Check if the PDF has at least one page
        if pdf_document.page_count > 0:
            # Load the first page (page number 0)
            page = pdf_document.load_page(0)
            # Extract text from the first page
            text = page.get_text()
        else:
            text = ""
        # Close the PDF document
        pdf_document.close()
        # Return the text in lowercase
        return text.lower()
    except Exception as e:
        print(f"Error reading first page of {pdf_path}: {e}")
        return ""


def process_text_based_on_patterns(text):
    # Define the patterns and corresponding functions
    patterns_functions = {
        r'Intermediary Name : Probus Insurance Broker Limited - BRR.*\bPPV\b': FUTURE_PC.extract_details,
        r'Intermediary Name : Probus Insurance Broker Limited - BRR.*\bPCV\b': FUTURE_PCV.extract_details,
        r'Intermediary Name : Probus Insurance Broker Limited - BRR.*\bGCV\b': FUTURE_GCV.extract_details
    }

    # Check for each pattern and call the corresponding function
    for pattern, func in patterns_functions.items():
        if re.search(pattern, text, re.MULTILINE | re.IGNORECASE):
            return func(text)

    return "No matching pattern found."


def main(pdf_path):
    # Extract text from the first page of the PDF
    text = extract_text_from_first_page(pdf_path)

    # Process the text based on patterns
    result = process_text_based_on_patterns(text)

    print(f"Processing result for {pdf_path}:")
    print(result)


if __name__ == "__main__":
    # Replace 'example.pdf' with the path to your PDF
    main(r"C:\Users\user\PycharmProjects\PDF_READER\PDF CONTAINER\FUTURE PDF\VD182278.pdf")
