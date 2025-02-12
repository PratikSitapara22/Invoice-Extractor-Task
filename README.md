# Invoice Extraction

## Overview
This project automates the extraction of invoice data from emails. It connects to an email inbox, retrieves invoices from attachments, extracts key information using OCR, and stores the data in MongoDB. The system supports **duplicate detection** and **recurring invoice classification** to streamline invoice management.

## Features
- **Email Integration:** Connects to Gmail or Outlook via IMAP.
- **Attachment Processing:** Extracts text from PDFs, PNGs, and JPGs.
- **OCR-Based Data Extraction:** Uses Tesseract to extract invoice details.
- **MongoDB Storage:** Saves extracted invoice data in a structured format.
- **Duplicate Invoice Detection:** Prevents duplicate invoices from being stored.
- **Recurring Invoice Classification:** Identifies subscription-based invoices and stores them separately.
- **User-Friendly Output:** Extracted data is printed in JSON format for easy readability.

## Installation & Setup

### Prerequisites
Ensure you have the following installed on your system:

- **Poppler (for PDF processing):** [Download Here](https://github.com/oschwartz10612/poppler-windows/releases)
- **Tesseract OCR:** [Download Here](https://github.com/UB-Mannheim/tesseract/wiki)
  - After installation, add the Tesseract and Poppler path to the environment variable manually or using the command prompt.
- **MongoDB:** [Download Here](https://www.mongodb.com/try/download/community)


### Clone the Repository
```bash
git clone https://github.com/PratikSitapara22/Invoice-Extractor-Task.git
cd Invoice-Extractor-Task
```

### Install Dependencies
```bash
pip install -r requirements.txt
```

## Usage

### Running the Script
```bash
python InvoiceExtractor.py
```

### Entering Email Credentials
1. **IMAP Server:** Enter `Gmail` or `Outlook`
2. **Email Address:** Enter your email address
3. **Password:** Enter your email password
   
   (OR use the ID-Password of Dummy Gmail provided via personal email on 12/02/2025)
4. **Filter Type:** Select one of the following:
   - `subject` (to filter emails by subject)
   - `sender` (to filter by email sender)
   - `attachments` (to filter emails with attachments only; requires no input)
5. **Filter Value:** Provide a keyword based on your Filter Type. Examples:
   - If `subject` is selected, enter a keyword like `invoice`.
   - If `sender` is selected, enter the email address (e.g., `abc@gmail.com`).
   - If `attachments` is selected, no input is needed.

### Checking Extracted Data in MongoDB
1. Open **MongoDB Compass** 
2. Navigate to **InvoiceDB** and check:
   - `Invoices` collection (for normal invoices)
   - `RecurringInvoices` collection (for subscription-based invoices)

### Testing with Sample Invoices

A `TestData` folder is included in this repository, containing **four sample invoices**:
- **2 Regular Invoices** (to test normal invoice extraction)
- **2 Subscription-Based Invoices** (to test recurring invoice classification)

These invoices are for **testing purposes only** and do not contain real data.

## Future Enhancements
For better results, **LLMs (Large Language Models)** or **Vision Models** can be utilized for extracting structured data from a large dataset of invoices using **orchestration frameworks** like **LangChain** and **Open-Source GenAI Models**. This could further improve accuracy and automate complex invoice processing.


