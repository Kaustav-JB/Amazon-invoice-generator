import os
import re
import pdfplumber
import pandas as pd

INPUT_FOLDER = "../Input folder"
OUTPUT_FILE = "../Output folder/extracted_invoice.xlsx"
os.makedirs("../Output folder", exist_ok=True)

def extract_field(text, pattern, group=1):
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return match.group(group).strip() if match and match.lastindex and match.lastindex >= group else ""

def extract_invoice_data(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    data = {}

    data["Order Number"] = extract_field(text, r"Order Number[:\s]*([0-9\-]+)")
    data["Invoice Number"] = extract_field(text, r"Invoice Number\s*[:]*\s*([A-Z0-9\-]+)")
    data["Order Date"] = extract_field(text, r"Order Date\s*[:]*\s*([\d\.\/]+)")
    data["Invoice Date"] = extract_field(text, r"Invoice Date\s*[:]*\s*([\d\.\/]+)")

    lines = text.splitlines()

    desc_line = ""
    for line in lines:
        if line.strip().startswith("1 "):
            desc_line = line.strip()
            break

    if desc_line:
        desc_clean = re.sub(r"1\s+", "", desc_line)
        desc_clean = re.sub(r"₹[\d,]+\.\d{2}.*", "", desc_clean)
        data["Item Description"] = " ".join(desc_clean.strip().split())
    else:
        data["Item Description"] = ""

    data["Unit Price"] = extract_field(text, r"₹([\d,]+\.\d{2})\s+1\s+₹[\d,]+\.\d{2}", group=1)
    data["Quantity"] = extract_field(text, r"₹[\d,]+\.\d{2}\s+(\d+)\s+₹[\d,]+\.\d{2}")
    data["Net Amount"] = extract_field(text, r"₹[\d,]+\.\d{2}\s+1\s+₹([\d,]+\.\d{2})", group=1)

    data["CGST"] = extract_field(text, r"CGST\s+₹([\d,]+\.\d{2})")
    data["SGST"] = extract_field(text, r"SGST\s+₹([\d,]+\.\d{2})")
    invoice_value = ""
    for i, line in enumerate(lines):
        if "Invoice Value" in line:
            for j in range(i, i + 3):
                if j < len(lines) and re.search(r"\d{1,3}(,\d{3})*\.\d{2}", lines[j]):
                    invoice_value = re.search(r"\d{1,3}(?:,\d{3})*\.\d{2}", lines[j]).group()
                    break
            break

    data["Invoice Value"] = invoice_value

    data["Seller Name"] = extract_field(text, r"Sold By\s*:\s*(.*?)\n\*")
    data["PAN"] = extract_field(text, r"PAN No\s*[:]*\s*([A-Z0-9]+)")
    data["GSTIN"] = extract_field(text, r"GST Registration No\s*[:]*\s*([A-Z0-9]+)")

    seller_name = ""
    for i, line in enumerate(lines):
        if "Sold By" in line:
            for j in range(i + 1, len(lines)):
                next_line = lines[j].strip()
                if next_line and not next_line.startswith("*"):
                    seller_name = next_line
                    break
            break
    data["Seller Name"] = seller_name

    data["Shipping Name"] = extract_field(text, r"Shipping Address\s*:\s*([A-Z ]+)")

    data["Source File"] = os.path.basename(pdf_path)
    return data

all_data = []
for filename in os.listdir(INPUT_FOLDER):
    if filename.lower().endswith(".pdf"):
        file_path = os.path.join(INPUT_FOLDER, filename)
        print(f"🔍Processing: {filename}")
        try:
            extracted = extract_invoice_data(file_path)
            all_data.append(extracted)
        except Exception as e:
            print(f"❌Failed to process {filename}: {e}")

if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅All extractions complete. Excel saved at: {OUTPUT_FILE}")
else:
    print("\n⚠️No invoices were successfully processed.")