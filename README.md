# 🧾 Amazon Invoice Data Extractor 

This project extracts structured data (like order details, seller info, taxes, and item descriptions) from Amazon invoice PDFs and exports it cleanly into an Excel sheet — ideal for reporting, reconciliation, and automation.

## 🚀 Features

- 🔍 Extracts from multiple PDF invoices in one go
- 🧠 Handles complex formatting & line-breaks intelligently
- 💰 Cleans amounts (removes ₹ symbols and commas)
- 📊 Exports clean, analysis-ready Excel file
- 💥 Handles edge cases (like multi-line fields, empty fields)

## 🔧 What It Extracts

- Order Number, Invoice Number, Order & Invoice Dates
- Item Description
- Unit Price, Quantity, Net Amount
- CGST, SGST, Total Tax, Total Amount
- Invoice Value
- Seller Name, GSTIN, PAN
- Billing & Shipping Names

## 🛠️ Installation

Clone the repo:
```
git clone https://github.com/Kaustav-JB/Amazon-invoice-generator.git
cd invoice-data-extractor
```
Install dependencies:
```
pip install -r requirements.txt
```

## 📥 Usage
 - Place your .pdf invoices in the ```Input folder/```
 - Run the script in the ```Coding folder/```:
```
python extract_invoice.py
```
 - Open the generated extracted_invoice.xlsx in ```Output folder/```

## 🙋‍♂️ Contributions & Support
Feel free to fork this project, improve it, and raise a PR.
Found a bug or need support? Open an issue.
