import os
import pandas as pd
from flask import Flask, render_template, request, send_file, make_response, jsonify
import xml.etree.ElementTree as ET
import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect
from werkzeug.utils import secure_filename

app = Flask(__name__)

# --- CONFIGURATION ---
UPLOAD_FOLDER = '/tmp/uploads' 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- 1. PDF PARSER (With Password Support) ---
def parse_pdf_to_dataframe(pdf_path, pdf_password=None):
    all_data = []
    try:
        # Password handling logic
        with pdfplumber.open(pdf_path, password=pdf_password or "") as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
                            if any(cleaned_row): 
                                all_data.append(cleaned_row)
    except PDFPasswordIncorrect:
        # Agar password galat hai ya chahiye, to ye error raise karo
        raise ValueError("PASSWORD_REQUIRED")
    except Exception as e:
        # Koi aur error ho (jaise wrong password format)
        if "password" in str(e).lower():
            raise ValueError("PASSWORD_REQUIRED")
        print(f"PDF Error: {e}")
        return pd.DataFrame()

    if not all_data: return pd.DataFrame()

    headers = all_data[0]
    data = all_data[1:]
    headers = [f"{col}_{i}" if headers.count(col) > 1 else col for i, col in enumerate(headers)]
    return pd.DataFrame(data, columns=headers)

# --- 2. XML GENERATOR ---
def generate_tally_xml(df, output_path, conversion_type, main_ledger_name):
    suspense_ledger_name = "Suspense Account"
    
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(req_desc, "REPORTNAME").text = "All Masters"
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    # Auto-Create Suspense Logic
    if conversion_type == 'bank':
        tally_msg_master = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        ledger = ET.SubElement(tally_msg_master, "LEDGER", {"NAME": suspense_ledger_name, "ACTION": "Create"})
        name_list = ET.SubElement(ledger, "NAME.LIST")
        ET.SubElement(name_list, "NAME").text = suspense_ledger_name
        ET.SubElement(ledger, "PARENT").text = "Suspense A/c" 
        ET.SubElement(ledger, "ISBILLWISEON").text = "No"
        ET.SubElement(ledger, "AFFECTSSTOCK").text = "No"

    # Vouchers Logic
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
        if 'file' not in request.files: 
            return jsonify({'error': "No file uploaded"}), 400
        
        file = request.files['file']
        if file.filename == '': 
            return jsonify({'error': "No file selected"}), 400

        conversion_type = request.form.get('type')
        main_ledger = request.form.get('main_ledger', 'Bank Account')
        pdf_password = request.form.get('password', None) # Password Frontend se aayega

        if file:
            filename = secure_filename(file.filename)
            base_name = os.path.splitext(filename)[0]
            xml_filename = f"{base_name}.xml"
            
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], xml_filename)
            
            file.save(input_path)
            
            df = pd.DataFrame()
            fname_lower = filename.lower()

            try:
                if fname_lower.endswith('.pdf'):
                    # Yahan Password Logic Call Hoga
                    df = parse_pdf_to_dataframe(input_path, pdf_password)
                    if df.empty: 
                         return jsonify({'error': "PDF seems empty or unrecognizable."}), 400
                
                elif fname_lower.endswith(('.xls', '.xlsx')):
                    engine = 'xlrd' if fname_lower.endswith('.xls') else 'openpyxl'
                    df = pd.read_excel(input_path, engine=engine)
                else:
                    return jsonify({'error': "Invalid format. Use PDF or Excel."}), 400

            except ValueError as e:
                # Agar Password Chahiye to Frontend ko batao
                if str(e) == "PASSWORD_REQUIRED":
                    return jsonify({'status': 'password_required'}), 401
                raise e

            df = df.fillna('')
            generate_tally_xml(df, output_path, conversion_type, main_ledger)
            
            return send_file(output_path, as_attachment=True, download_name=xml_filename)
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)