import os
import pandas as pd
from flask import Flask, render_template, request, send_file, make_response
import xml.etree.ElementTree as ET
import pdfplumber
from werkzeug.utils import secure_filename

app = Flask(__name__)

# --- CONFIGURATION ---
UPLOAD_FOLDER = '/tmp/uploads' 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- 1. PDF PARSER ---
def parse_pdf_to_dataframe(pdf_path):
    all_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
                            if any(cleaned_row): 
                                all_data.append(cleaned_row)
    except Exception as e:
        print(f"PDF Error: {e}")
        return pd.DataFrame()

    if not all_data: return pd.DataFrame()

    headers = all_data[0]
    data = all_data[1:]
    headers = [f"{col}_{i}" if headers.count(col) > 1 else col for i, col in enumerate(headers)]
    return pd.DataFrame(data, columns=headers)

# --- 2. XML GENERATOR ---
def generate_tally_xml(df, output_path, conversion_type, main_ledger_name, suspense_ledger_name):
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(req_desc, "REPORTNAME").text = "Vouchers"
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    df.columns = [str(c).strip() for c in df.columns]
    
    for index, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        
        date_val = ""
        for col in df.columns:
            if "date" in col.lower():
                date_val = row[col]
                break
        try:
            date_obj = pd.to_datetime(date_val, dayfirst=True)
            tally_date = date_obj.strftime('%Y%m%d')
        except:
            tally_date = "20240401"

        amount = 0
        debit = 0
        credit = 0
        for col in df.columns:
            c_low = col.lower()
            val = str(row[col]).replace(',', '').strip()
            if not val: continue
            try:
                if "debit" in c_low or "dr" in c_low or "withdrawal" in c_low: debit = float(val)
                elif "credit" in c_low or "cr" in c_low or "deposit" in c_low: credit = float(val)
                elif "amount" in c_low: amount = float(val)
            except: pass

        if debit > 0: amount = debit
        elif credit > 0: amount = credit
        
        narration = ""
        party_name = ""
        for col in df.columns:
            if "party" in col.lower() or "name" in col.lower():
                party_name = str(row[col])
                break
        for col in df.columns:
            if "particular" in col.lower() or "narration" in col.lower() or "description" in col.lower():
                narration = str(row[col])
                if not party_name: party_name = narration
                break

        if conversion_type == 'sales':
            vch_type = "Sales"
            is_party_pos = "Yes"
            is_main_pos = "No"
        elif conversion_type == 'purchase':
            vch_type = "Purchase"
            is_party_pos = "No"
            is_main_pos = "Yes"
        else:
            if debit > 0:
                vch_type = "Payment"
                is_party_debit = "Yes"
                is_bank_credit = "No"
            else:
                vch_type = "Receipt"
                is_party_debit = "No"
                is_bank_credit = "Yes"
            
            voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
            ET.SubElement(voucher, "DATE").text = tally_date
            ET.SubElement(voucher, "NARRATION").text = narration
            ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
            
            l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l1, "LEDGERNAME").text = suspense_ledger_name
            ET.SubElement(l1, "ISDEEMEDPOSITIVE").text = is_party_debit
            ET.SubElement(l1, "AMOUNT").text = str(-amount if is_party_debit == "Yes" else amount)

            l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
            ET.SubElement(l2, "ISDEEMEDPOSITIVE").text = is_bank_credit
            ET.SubElement(l2, "AMOUNT").text = str(-amount if is_bank_credit == "Yes" else amount)
            continue 

        voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
        ET.SubElement(voucher, "DATE").text = tally_date
        ET.SubElement(voucher, "NARRATION").text = narration
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
        
        l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(l1, "LEDGERNAME").text = party_name
        ET.SubElement(l1, "ISDEEMEDPOSITIVE").text = is_party_pos
        ET.SubElement(l1, "AMOUNT").text = str(-amount if is_party_pos == "Yes" else amount)

        l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
        ET.SubElement(l2, "ISDEEMEDPOSITIVE").text = is_main_pos
        ET.SubElement(l2, "AMOUNT").text = str(-amount if is_main_pos == "Yes" else amount)

    tree = ET.ElementTree(envelope)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

# --- 3. ROUTES ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files: return "No file uploaded"
        file = request.files['file']
        if file.filename == '': return "No file selected"

        conversion_type = request.form.get('type')
        main_ledger = request.form.get('main_ledger', 'Sales Account')
        suspense_ledger = request.form.get('suspense_ledger', 'Suspense Account')

        if file:
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], "Tally_Import.xml")
            
            file.save(input_path)
            
            df = pd.DataFrame()
            fname_lower = filename.lower()

            if fname_lower.endswith('.pdf'):
                df = parse_pdf_to_dataframe(input_path)
                if df.empty: return "<h1>Error:</h1><p>PDF Error. Try converting to Excel.</p>"
            elif fname_lower.endswith(('.xls', '.xlsx')):
                engine = 'xlrd' if fname_lower.endswith('.xls') else 'openpyxl'
                df = pd.read_excel(input_path, engine=engine)
            else:
                return "Format Error: Use PDF or Excel."

            df = df.fillna('')
            generate_tally_xml(df, output_path, conversion_type, main_ledger, suspense_ledger)
            
            # --- MAGIC FIX: Cookie Set Karna ---
            response = make_response(send_file(output_path, as_attachment=True))
            # Yeh cookie batayegi ki download complete ho gaya
            response.set_cookie('file_download_token', 'done', max_age=60, path='/')
            return response
            
    except Exception as e:
        return f"<h1>Error:</h1><p>{str(e)}</p>"

if __name__ == '__main__':
    app.run(debug=True)