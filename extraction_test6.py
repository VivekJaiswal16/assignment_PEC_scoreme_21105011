import pdfplumber
from pytesseract import image_to_string
from PIL import Image
import pandas as pd
from openpyxl import Workbook

def extract_text_from_scanned_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages):
            print(f"Processing page {page_number + 1}...")
            # Extract the page as an image
            img = page.to_image(resolution=300).original  # Increase resolution for better OCR
            # Perform OCR on the image
            page_text = image_to_string(img)
            text += page_text + "\n"  # Add a newline between pages
    return text

def parse_text_to_table(text):
    # Split the text into lines
    lines = text.split("\n")
    
    # Initialize a list to store rows
    table = []
    
    # Process each line
    for line in lines:
        # Split the line into columns (adjust the delimiter as needed)
        columns = line.split()  # Split by spaces (default)
        if columns:  # Ignore empty lines
            table.append(columns)
    
    return table

def save_table_to_excel(table, output_excel_path):
    # Convert the table to a DataFrame
    df = pd.DataFrame(table)
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False, header=False)
    print(f"Table extracted and saved to {output_excel_path}")

def main(pdf_path, output_excel_path):
    # Extract text from the scanned PDF
    text = extract_text_from_scanned_pdf(pdf_path)
    
    # Parse the text into a table
    table = parse_text_to_table(text)
    
    # Save the table to an Excel file
    save_table_to_excel(table, output_excel_path)

if __name__ == "__main__":
    pdf_path = "test6.pdf"  # Replace with your PDF file path
    output_excel_path = "output_test6.xlsx"  # Replace with your desired output Excel file path
    main(pdf_path, output_excel_path)