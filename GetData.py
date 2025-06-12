import pdfplumber
import pandas as pd
import re

def extract_invoice_data(pdf_path):
    data = {}
    line_items = []
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

        # Invoice Number
        m = re.search(r"Invoice (Number|No)[\s:#]*([A-Za-z0-9\-]+)", text)
        data['Invoice Number'] = m.group(2).strip() if m else ""

        # Order ID
        m = re.search(r"Order (ID|Number)[\s:#]*([A-Za-z0-9\-]+)", text)
        data['Order ID'] = m.group(2).strip() if m else ""

        # Invoice Date
        m = re.search(r"Invoice Date[\s:#]*([0-9\-./]+)", text)
        data['Invoice Date'] = m.group(1).strip() if m else ""

        # Order Date
        m = re.search(r"Order Date[\s:#]*([0-9\-./, :AMP]+)", text)
        data['Order Date'] = m.group(1).strip() if m else ""

        # Seller
        m = re.search(r"Sold By[\s:]*([^\n]+)", text)
        data['Seller'] = m.group(1).strip() if m else ""

        # GSTIN
        m = re.search(r"GSTIN[\s:]*([A-Z0-9]+)", text)
        data['GSTIN'] = m.group(1).strip() if m else ""

        # PAN
        m = re.search(r"PAN[\s:]*([A-Z0-9]+)", text)
        data['PAN'] = m.group(1).strip() if m else ""

        # Billing Address (grab lines after 'Billing Address' up to next blank line)
        m = re.search(r"Billing Address[\s:]*([\s\S]+?)\n\n", text)
        data['Billing Address'] = m.group(1).strip().replace('\n', ', ') if m else ""

        # Shipping Address
        m = re.search(r"Shipping Address[\s:]*([\s\S]+?)\n\n", text)
        data['Shipping Address'] = m.group(1).strip().replace('\n', ', ') if m else ""

        # Try to extract tables from all pages and concatenate
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and any("Description" in str(cell) or "Product" in str(cell) for cell in table[0]):
                    headers = [h.strip() for h in table[0]]
                    for row in table[1:]:
                        if any(cell.strip() for cell in row):  
                            item = dict(zip(headers, row))
                            line_items.append(item)

        m = re.search(r"TOTAL PRICE[\s:₹]*([0-9.,]+)", text)
        data['Total Price'] = m.group(1).strip() if m else ""

        m = re.search(r"TOTAL QTY[\s:]*([0-9]+)", text)
        data['Total Qty'] = m.group(1).strip() if m else ""

    return data, line_items

def save_to_excel(header_data, line_items, output_path):
    # Write header data to first sheet
    df_header = pd.DataFrame([header_data])

    # Write items to second sheet
    df_items = pd.DataFrame(line_items)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_header.to_excel(writer, sheet_name='Invoice Header', index=False)
        df_items.to_excel(writer, sheet_name='Line Items', index=False)

if __name__ == "__main__":
    input_pdf = "invoice4.pdf"  # PDF file
    output_excel = "extracted_invoice4.xlsx"

    header_data, line_items = extract_invoice_data(input_pdf)
    save_to_excel(header_data, line_items, output_excel)
    print("Extraction complete. Data saved to", output_excel)
