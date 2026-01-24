import os
import pandas as pd
from flask import Flask, render_template, request, send_file
import xml.etree.ElementTree as ET
import pdfplumber  # PDF library

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- 1. PDF HANDLING FUNCTION ---
def parse_pdf_to_dataframe(pdf_path):
    all_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            # Safai: None values ko empty string banao
                            cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
                            # Khali rows ko ignore karo
                            if any(cleaned_row): 
                                all_data.append(cleaned_row)
    except Exception as e:
        print(f"PDF Error: {e}")
        return pd.DataFrame() # Return empty if fail

    if not all_data:
        return pd.DataFrame()

    # Pehli row ko Header maano
    headers = all_data[0]
    data = all_data[1:]
    
    # Headers unique hone chahiye (Duplicate columns fix)
    headers = [f"{col}_{i}" if headers.count(col) > 1 else col for i, col in enumerate(headers)]

    return pd.DataFrame(data, columns=headers)

# --- 2. XML GENERATOR FUNCTION ---
def generate_tally_xml(df, output_path, conversion_type, main_ledger_name, suspense_ledger_name):
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(req_desc, "REPORTNAME").text = "Vouchers"
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    # Column names ko lowercase aur strip karo taaki matching aasan ho
    df.columns = [str(c).strip() for c in df.columns]
    
    for index, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        
        # --- SMART COLUMN MAPPING ---
        # Date dhoondo
        date_val = ""
        for col in df.columns:
            if "date" in col.lower():
                date_val = row[col]
                break
        
        try:
            date_obj = pd.to_datetime(date_val, dayfirst=True)
            tally_date = date_obj.strftime('%Y%m%d')
        except:
            tally_date = "20240401" # Default Date

        # Amount Logic (Debit/Credit/Amount columns)
        amount = 0
        debit = 0
        credit = 0
        
        for col in df.columns:
            c_low = col.lower()
            val = str(row[col]).replace(',', '').strip()
            if not val: continue
            
            try:
                if "debit" in c_low or "dr" in c_low or "withdrawal" in c_low:
                    debit = float(val)
                elif "credit" in c_low or "cr" in c_low or "deposit" in c_low:
                    credit = float(val)
                elif "amount" in c_low:
                    amount = float(val)
            except:
                pass

        if debit > 0: amount = debit
        elif credit > 0: amount = credit
        
        # Narration / Party Name Logic
        narration = ""
        party_name = ""
        
        # Pehle Party column dhoondo
        for col in df.columns:
            if "party" in col.lower() or "name" in col.lower():
                party_name = str(row[col])
                break
        
        # Agar Party nahi mila, to Narration/Particulars dhoondo
        for col in df.columns:
            if "particular" in col.lower() or "description" in col.lower() or "narration" in col.lower():
                narration = str(row[col])
                if not party_name: party_name = narration # Fallback
                break

        # --- XML CREATION ---
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

        else: # Bank
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

# --- 3. ROUTES ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files: return "No file uploaded"
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
            filename = file.filename.lower()
            df = pd.DataFrame()

            # --- FILE FORMAT CHECK ---
            if filename.endswith('.pdf'):
                df = parse_pdf_to_dataframe(input_path)
                if df.empty:
                    return "Error: PDF mein koi table nahi mila. Kripya Excel convert karke try karein."
            
            elif filename.endswith(('.xls', '.xlsx')):
                # Old vs New Excel Logic
                if filename.endswith('.xls'):
                    df = pd.read_excel(input_path, engine='xlrd')
                else:
                    df = pd.read_excel(input_path, engine='openpyxl')
            else:
                return "Error: Invalid file format. Only PDF, XLS, XLSX supported."

            df = df.fillna('')
            
            # Generate XML
            generate_tally_xml(df, output_path, conversion_type, main_ledger, suspense_ledger)
            
            return send_file(output_path, as_attachment=True)
            
        except Exception as e:
            # Yeh error screen par dikhega instead of crashing
            return f"System Error: {str(e)} <br><br> Tip: File format check karein ya Excel use karein."

if __name__ == '__main__':
    app.run(debug=True)