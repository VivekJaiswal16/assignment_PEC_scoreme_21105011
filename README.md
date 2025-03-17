# Bank Statement Extractor

This project extracts transaction data from bank statements in **PDF format** using **OCR (Tesseract)** and **pdfplumber**, then saves the structured data to an **Excel file** via a **Tkinter GUI**.

## Installation

To install the required dependencies, run the following command:

```bash
pip install pdfplumber pandas pytesseract pillow tk
```

### Additional Dependencies:
- **`pdfplumber`**: Extracts text from PDFs.
- **`pandas`**: Used for data processing and saving to Excel.
- **`re`**: Built-in module for regex (no installation needed).
- **`tkinter`**: Built-in Python GUI module (no installation needed for standard Python distributions).
- **`pytesseract`**: OCR tool to extract text from images (requires Tesseract installation).
- **`pillow`**: Required for `pytesseract` to process images.

## Installing Tesseract OCR

If you haven't installed **Tesseract OCR**, follow these steps:

### Windows:
1. Download and install Tesseract from:
   [Tesseract Installation Guide](https://github.com/UB-Mannheim/tesseract/wiki)
2. Add the installation path to system environment variables (if needed).

### Linux (Ubuntu/Debian):
```bash
sudo apt install tesseract-ocr
```

### macOS (Homebrew):
```bash
brew install tesseract
```

## Usage
1. Run the script to open the GUI.
2. Click **"Upload PDF"** to select a bank statement.
3. Click **"Download Excel"** to save the extracted transactions.

## Features
✅ Extracts transactions from PDF bank statements using OCR.
✅ Converts extracted data into a structured Excel file.
✅ Provides an easy-to-use graphical interface with Tkinter.


