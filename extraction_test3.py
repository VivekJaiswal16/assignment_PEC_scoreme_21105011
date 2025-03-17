import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from pytesseract import image_to_string

# Function to extract text using OCR
def extract_text_with_ocr(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            img = page.to_image(resolution=400).original
            text += image_to_string(
                img, 
                config='--psm 6 -c preserve_interword_spaces=1'
            ) + "\n"
    return text

# Function to parse transaction data from extracted text
def parse_transaction_data(text):
    """Parse OCR text into structured transaction data"""
    transactions = []
    current_date = None
    current_description = []
    current_amount = None
    current_balance = None
    
    for line in text.split('\n'):
        line = line.strip()
        
        # Skip header/footer lines
        if any(x in line for x in ["BANK NAME", "BRANCH NAME", "IFSC Code", "***END"]):
            continue
            
        # Match transaction lines (date, description, amount, balance)
        if len(line) > 10 and line[2] == '-' and line[6] == '-':  # Date format: dd-mmm-yyyy
            if current_date:
                # Save the previous transaction
                transactions.append([
                    current_date,
                    ' '.join(current_description).strip(),
                    current_amount,
                    current_balance
                ])
            
            # Start a new transaction
            parts = line.split(maxsplit=2)
            current_date = parts[0]
            current_description = [parts[1]] if len(parts) > 1 else []
            if len(parts) > 2:
                # Extract amount and balance from the remaining part
                amount_balance = parts[2].rsplit(maxsplit=1)
                if len(amount_balance) == 2:
                    current_amount = amount_balance[0]
                    current_balance = amount_balance[1]
                else:
                    current_amount = amount_balance[0]
                    current_balance = None
            else:
                current_amount = None
                current_balance = None
        
        elif current_date:
            # Continuation of description
            current_description.append(line)
    
    # Save the last transaction
    if current_date:
        transactions.append([
            current_date,
            ' '.join(current_description).strip(),
            current_amount,
            current_balance
        ])
    
    return transactions

# Function to save extracted data to Excel
def save_to_excel(data, output_path):
    df = pd.DataFrame(data, columns=["Date", "Description", "Amount", "Balance"])
    numeric_cols = ["Amount", "Balance"]
    for col in numeric_cols:
        df[col] = df[col].replace('', pd.NA).str.replace(r'[^\d.]', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df.dropna(subset=numeric_cols, how='all', inplace=True)
    df.to_excel(output_path, index=False)
    print(f"Successfully saved {len(df)} transactions to {output_path}")

# Function to upload file
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        raw_text = extract_text_with_ocr(file_path)
        transactions = parse_transaction_data(raw_text)
        if not transactions:
            messagebox.showerror("Error", "No transactions found in the PDF.")
        else:
            global extracted_data
            extracted_data = transactions
            messagebox.showinfo("Success", "File uploaded and transactions extracted!")

# Function to save extracted data to Excel
def save_to_excel_gui():
    if not extracted_data:
        messagebox.showerror("Error", "No data to save. Please upload a valid PDF first.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        save_to_excel(extracted_data, file_path)
        messagebox.showinfo("Success", f"Data saved to {file_path}")

# Initialize the Tkinter window
root = tk.Tk()
root.title("Bank Statement Extractor")
root.geometry("400x200")

tk.Label(root, text="Upload a bank statement PDF").pack(pady=10)

upload_btn = tk.Button(root, text="Upload PDF", command=upload_file)
upload_btn.pack(pady=5)

download_btn = tk.Button(root, text="Download Excel", command=save_to_excel_gui)
download_btn.pack(pady=5)

extracted_data = []

root.mainloop()