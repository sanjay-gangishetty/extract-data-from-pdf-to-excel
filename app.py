import pdfplumber
import pandas as pd
import re

def extract_data_from_pdf(pdf_path):
    data = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            print("Extracted Text:\n", text)  # Print the text for debugging
            if not text:
                continue
            # Use regular expressions to find the fields
            invoice_number_match = re.search(r"Invoice Number:\s*(\S+)", text)
            date_match = re.search(r"Date:\s*([A-Za-z]+\s\d{1,2},\s\d{4})", text)
            total_match = re.search(r"Total:\s*\$?(\S+)", text)
            
            if invoice_number_match:
                data['Invoice Number'] = invoice_number_match.group(1)
            if date_match:
                data['Date'] = date_match.group(1)
            if total_match:
                data['Total'] = total_match.group(1)
    return data

def update_excel_with_data(excel_path, data):
    df = pd.read_excel(excel_path)
    print(data),
    # Assuming the Excel file has columns 'Invoice Number', 'Date', 'Total', etc.
    new_row = {
        'Invoice Number': data.get('Invoice Number', ''),
        'Date': data.get('Date', ''),
        'Total': data.get('Total', ''),
    }
    
    df = df.append(new_row, ignore_index=True)
    df.to_excel(excel_path, index=False)
    print("Excel file updated successfully!")

# Test the process
def test_process(pdf_path, excel_path):
    data = extract_data_from_pdf(pdf_path)
    print("Extracted Data:", data)
    if not data:
        print("No data extracted from PDF.")
        return
    update_excel_with_data(excel_path, data)

# Define the paths to your PDF and Excel files
pdf_path = 'testPdf.pdf'  # Replace with your actual PDF file path
excel_path = 'test.xlsx'     # Replace with your actual Excel file path

# Run the test
test_process(pdf_path, excel_path)
