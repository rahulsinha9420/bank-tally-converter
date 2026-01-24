import os
import pandas as pd
from flask import Flask, render_template, request, send_file
import xml.etree.ElementTree as ET
import pdfplumber  # PDF padhne ke liye

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- PDF PARSER FUNCTION ---
def parse_pdf_to_dataframe(pdf_path):
    all_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Page se tables nikalo
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Khali rows hatao aur data saaf karo
                    cleaned_row = [str(cell).strip() if cell else '' for cell in row]
                    # Check karo row khali to nahi
                    if any(cleaned_row):
                        all_data.append(cleaned_row)
    
    if not all_data:
        return pd.DataFrame()

    # Pehli row ko Header man lete hain
    headers = all_data[0]
    data = all_data[1:]
    
    # DataFrame banao
    df = pd.DataFrame(data, columns=headers)
    return df

# --- XML GENERATOR ---
def generate_tally_xml(df, output_path, conversion_type, main_ledger_name, suspense_ledger_name):
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(req_desc, "REPORTNAME").text = "Vouchers"
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    # Column Names ko Standardize karo (Case Insensitive)
    df.columns = [c.strip() for c in df.columns]
    
    for index, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        
        # --- SMART DATA MAPPING ---
        # Date dhundo (Date, txn date, value date etc.)
        date_val = ""
        for col in df.columns:
            if "date" in col.lower():
                date_val = row[col]
                break
        
        try:
            date_obj = pd.to_datetime(date_val, dayfirst=True)
            tally_date = date_obj.strftime('%Y%m%d')
        except:
            tally_date = "20240401" # Default agar date na mile

        # Amount Dhundo
        amount = 0
        debit = 0
        credit = 0
        
        # Agar 'Debit' aur 'Credit' columns alag hain
        for col in df.columns:
            if "debit" in col.lower() or "dr" in col.lower():
                try: debit = float(str(row[col]).replace(',', '').replace('Dr', '').strip())
                except: pass
            if "credit" in col.lower() or "cr" in col.lower():
                try: credit = float(str(row[col]).replace(',', '').replace('Cr', '').strip())
                except: pass
                
        if debit > 0: amount = debit
        elif credit > 0: amount = credit
        else:
            # Agar sirf 'Amount' column hai
            for col in df.columns:
                if "amount" in col.lower():
                    try: amount = float(str(row[col]).replace(',', ''))
                    except: pass
                    break

        # Narration / Party
        narration = ""
        party_name = ""
        for col in df.columns:
            if "particular" in col.lower() or "description" in col.lower() or "narration" in col.lower():
                narration = str(row[col])
                party_name = str(row[col]) # Default Party Name = Narration
                break
        
        # Agar Party column alag se hai
        for col in df.columns:
            if "party" in col.lower():
                party_name = str(row[col])
                break

        # XML Structure (Wahi purana logic)
        if conversion_type == 'sales':
            vch_type = "Sales"
            voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
            ET.SubElement(voucher, "DATE").text = tally_date
            ET.SubElement(voucher, "NARRATION").text = narration
            ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
            
            l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l1, "LEDGERNAME").text = party_name
            ET.SubElement(l1, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(l1, "AMOUNT").text = str(-amount)

            l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
            ET.SubElement(l2, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(l2, "AMOUNT").text = str(amount)

        elif conversion_type == 'purchase':
            vch_type = "Purchase"
            voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
            ET.SubElement(voucher, "DATE").text = tally_date
            ET.SubElement(voucher, "NARRATION").text = narration
            ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
            
            l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l1, "LEDGERNAME").text = party_name
            ET.SubElement(l1, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(l1, "AMOUNT").text = str(amount) 

            l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
            ET.SubElement(l2, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(l2, "AMOUNT").text = str(-amount)

        else:
            # Bank Logic
            if debit > 0:
                vch_type = "Payment"
                is_party_debit = "Yes"
                is_bank_credit = "No"
                final_amt = debit
            else:
                vch_type = "Receipt"
                is_party_debit = "No"
                is_bank_credit = "Yes"
                final_amt = credit

            voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
            ET.SubElement(voucher, "DATE").text = tally_date
            ET.SubElement(voucher, "NARRATION").text = narration
            ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type

            l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l1, "LEDGERNAME").text = suspense_ledger_name
            ET.SubElement(l1, "ISDEEMEDPOSITIVE").text = is_party_debit
            amt1 = -final_amt if is_party_debit == "Yes" else final_amt
            ET.SubElement(l1, "AMOUNT").text = str(amt1)

            l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
            ET.SubElement(l2, "ISDEEMEDPOSITIVE").text = is_bank_credit
            amt2 = -final_amt if is_bank_credit == "Yes" else final_amt
            ET.SubElement(l2, "AMOUNT").text = str(amt2)

    tree = ET.ElementTree(envelope)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files: return "No file"
    file = request.files['file']
    if file.filename == '': return "No file selected"

    conversion_type = request.form.get('type')
    main_ledger = request.form.get('main_ledger', 'Sales Account')
    suspense_ledger = request.form.get('suspense_ledger', 'Suspense Account')

    if file:
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], "Tally_Import.xml")
        file.save(input_path)
        
        try:
            # --- FILE FORMAT DETECTION ---
            filename = file.filename.lower()
            
            if filename.endswith('.pdf'):
                # Agar PDF hai to PDF parser use karo
                df = parse_pdf_to_dataframe(input_path)
            elif filename.endswith(('.xls', '.xlsx')):
                # Agar Excel hai
                engine = 'xlrd' if filename.endswith('.xls') else 'openpyxl'
                df = pd.read_excel(input_path, engine=engine)
            else:
                return "Unsupported file format. Please upload PDF or Excel."

            df = df.fillna('')
            
            # Agar data khali hai
            if df.empty:
                return "Error: Could not extract table from this PDF. Try converting it to Excel first."

            generate_tally_xml(df, output_path, conversion_type, main_ledger, suspense_ledger)
            
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"Error: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)