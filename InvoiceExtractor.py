"""
Invoice Processing System

This script connects to an email account, extracts invoices from email attachments, performs OCR for text extraction, and 
stores structured data in a MongoDB database.The system supports duplicate detection and recurring invoice classification.
""" 

import os
import imaplib
import email
import pdf2image
import pytesseract
import pymongo
import re
import json
from email.header import decode_header
from PIL import Image
from bson import ObjectId

# Connection Setup for MongoDB
def connect_to_mongo():
    """Connects to MongoDB and returns the collections for invoices and recurring invoices."""
    try:
        client = pymongo.MongoClient("mongodb://localhost:27017/", serverSelectionTimeoutMS=5000)
        client.server_info()  # Check if MongoDB is reachable
        db = client["InvoiceDB"]
        collection = db["Invoices"]
        recurring_collection = db["RecurringInvoices"]
        return collection, recurring_collection
    except pymongo.errors.ServerSelectionTimeoutError:
        print("Error: Unable to connect to MongoDB. Ensure MongoDB is running.")
        return None, None

# Connecting to the email server
def connect_to_email(imap_server, email_user, email_pass):
    """
    Connects to the email server using the given credentials and selects the inbox.
        
    Args:
        imap_server (str): IMAP server address (e.g., imap.gmail.com).
        email_user (str): Email address.
        email_pass (str): Email password.
    
    Returns:
        imaplib.IMAP4_SSL: IMAP connection object if successful, otherwise None.
    """
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_pass)
        mail.select("inbox")
        print(f"Successfully connected to: {'Gmail' if 'gmail' in imap_server else 'Outlook'}")
        return mail
    except imaplib.IMAP4.error:
        print("Error: Authentication failed. Check email and password.")
        return None
    except Exception as e:
        print("Error connecting to email:", e)
        return None

# Searching emails based on filter
def search_emails(mail, filter_type, filter_value):
    """
    Searches for emails based on a filter type (subject, sender, or attachments).
    
    Args:
        mail (imaplib.IMAP4_SSL): IMAP connection object.
        filter_type (str): Type of filter ('subject', 'sender', 'attachments').
        filter_value (str): Value for the filter.
    
    Returns:
        list: List of email UIDs matching the search criteria.
    """
    search_filters = {
        "subject": f'(SUBJECT "{filter_value}")',
        "sender": f'(FROM "{filter_value}")',
        "attachments": "(X-GM-RAW \"has:attachment\")"
    }
    
    if filter_type not in search_filters:
        print("Invalid filter type!")
        return []
    
    result, data = mail.search(None, search_filters[filter_type])
    if result != "OK":
        print("Error searching emails.")
        return []
    return data[0].split()

# Processing email and extracting attachments
def process_email(mail, email_uid, save_folder="Invoices"):
    """
    Processes an email to extract metadata and download attachments.
    
    Args:
        mail (imaplib.IMAP4_SSL): IMAP connection object.
        email_uid (bytes): UID of the email to process.
        save_folder (str): Directory to save attachments.
    
    Returns:
        dict: Extracted metadata.
    """
    try:
        os.makedirs(save_folder, exist_ok=True)
        result, email_data = mail.fetch(email_uid, "(RFC822)")
        if result != "OK":
            print(f"Failed to fetch email UID: {email_uid}")
            return None
        
        msg = email.message_from_bytes(email_data[0][1])
        sender = msg.get("From", "Unknown")
        subject, encoding = decode_header(msg.get("Subject", "Unknown"))[0]
        subject = subject.decode(encoding) if isinstance(subject, bytes) else subject
        
        attachments = []
        for part in msg.walk():
            if part.get("Content-Disposition"):
                filename = part.get_filename()
                if filename and filename.lower().endswith((".pdf", ".jpg", ".png")):
                    filepath = os.path.join(save_folder, filename)
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    attachments.append(filepath)

        return {"email_uid": str(email_uid), "sender": sender, "subject": subject, "attachments": attachments} if attachments else None
    
    except Exception as e:
        print(f"Error processing email UID {email_uid}: {e}")
        return None

# Extracting text from attachments
def extract_text_from_attachments(attachments):
    """Extracts text from invoice attachments using OCR."""
    extracted_texts = {}
    for file in attachments:
        if file.lower().endswith(".pdf"):
            images = pdf2image.convert_from_path(file)
            text = "\n".join([pytesseract.image_to_string(img).strip() for img in images])
        elif file.lower().endswith((".jpg", ".png")):
            text = pytesseract.image_to_string(Image.open(file))
        else:
            continue
        extracted_texts[file] = text
    return extracted_texts

# Extracting structured invoice data
def extract_invoice_data(text):
    """Extracts key invoice data such as invoice number, amount, due date, and payment status from text using regular expression (re)."""
    invoice_details = {}
    invoice_details["invoice_number"] = re.search(r"([A-Z]{3,5}\d{6,8})", text, re.IGNORECASE).group(1) if re.search(r"([A-Z]{3,5}\d{6,8})", text, re.IGNORECASE) else "Unknown"
    invoice_details["amount"] = re.search(r"€\s?([\d,]+\.\d{2})", text).group(1) if re.search(r"€\s?([\d,]+\.\d{2})", text) else "Unknown"
    invoice_details["due_date"] = re.search(r"Due Date[:\s]*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE).group(1) if re.search(r"Due Date[:\s]*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE) else "Unknown"
    invoice_details["payment_status"] = "Paid" if re.search(r"\bPaid\b", text, re.IGNORECASE) else "Unpaid"
    return invoice_details

# Storing extracted data in MongoDB and printing structured JSON output
def store_in_mongo(collection, recurring_collection, extracted_data):
    """Stores invoice data in MongoDB"""
    if collection.find_one({"email_uid": extracted_data["email_uid"], "sender": extracted_data["sender"], "invoice_number": extracted_data["invoice_number"]}):
        print("Duplicate invoice detected. Skipping storage.")
        return
    
    if collection.find_one({"sender": extracted_data["sender"], "amount": extracted_data["amount"]}) and extracted_data["due_date"] != "Unknown":
        recurring_collection.insert_one(extracted_data)
        print("Stored in RecurringInvoices.")
    else:
        collection.insert_one(extracted_data)
        print("Stored in Invoices.")

# Main Execution
def main():
    imap_server = "imap.gmail.com" if "gmail" in input("Enter IMAP Server (Gmail/Outlook): ").lower() else "outlook.office365.com"
    email_user = input("Enter Email Address: ")
    email_pass = input("Enter Email Password: ")
    filter_type = input("Search by (subject/sender/attachments): ")
    filter_value = input("Enter Filter Value: ") if filter_type != "attachments" else ""
    
    mail = connect_to_email(imap_server, email_user, email_pass)
    if not mail:
        return
    
    collection, recurring_collection = connect_to_mongo()
    if collection is None:
        return
    
    email_uids = search_emails(mail, filter_type, filter_value)
    extracted_invoices = []
    
    for uid in email_uids:
        email_data = process_email(mail, uid)
        if not email_data:
            continue
        
        extracted_texts = extract_text_from_attachments(email_data["attachments"])
        invoice_details = extract_invoice_data(" ".join([text.decode('utf-8', 'ignore') if isinstance(text, bytes) else text for text in extracted_texts.values()]))
        email_data.update(invoice_details)
        store_in_mongo(collection, recurring_collection, email_data)
        extracted_invoices.append(email_data)


    # printing the final structured JSON output.
    print(f"Extracted ALL Invoice Data: {json.dumps(extracted_invoices, default=str, indent=4)}")
    mail.logout()
    print("************Successfully extracted and stored invoice data from your email. Thank You!************")

if __name__ == "__main__":
    main()
