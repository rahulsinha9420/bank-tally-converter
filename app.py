import os
import io
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
import xml.etree.ElementTree as ET
import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "BANKFLOW_SINGLE_PRO_RAHUL"

# --- CONFIGURATION ---
UPLOAD_FOLDER = '/tmp/uploads' 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- 1. SMART PDF PARSER ---
def parse_pdf_to_dataframe(pdf_path, pdf_password=None):
    all_data = []
    try:
        pwd = pdf_password if pdf_password else ""
        with pdfplumber.open(pdf_path, password=pwd) as pdf:
            for page in pdf.pages:
                try:
                    tables = page.extract_tables() 
                    if tables:
                        for table in tables:
                            for row in table:
                                cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
                                if any(cleaned_row): all_data.append(cleaned_row)
                except Exception: continue 
    except PDFPasswordIncorrect: raise ValueError("PASSWORD_REQUIRED")
    except Exception as e:
        if "password" in str(e).lower(): raise ValueError("PASSWORD_REQUIRED")
        return pd.DataFrame()

    if not all_data: return pd.DataFrame()

    # Smart Header Detection (IDFC & Multiple Bank Fix)
    header_index = 0
    max_score = 0
    keywords = ['date', 'particulars', 'description', 'narration', 'debit', 'credit', 'withdrawal', 'deposit', 'balance', 'val date', 'txn date']
    
    for i, row in enumerate(all_data[:25]):
        row_str = " ".join([str(x).lower() for x in row])
        score = sum(1 for kw in keywords if kw in row_str)
        if score > max_score:
            max_score = score
            header_index = i
            
    headers = all_data[header_index]
    data = all_data[header_index+1:]
    headers = [f"{col}_{i}" if headers.count(col) > 1 else col for i, col in enumerate(headers)]
    return pd.DataFrame(data, columns=headers)

# --- 2. XML GENERATOR ---
def generate_tally_xml(df, main_ledger_name):
    suspense = "Suspense Account"
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    tally_msg_m = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
    ledger = ET.SubElement(tally_msg_m, "LEDGER", {"NAME": suspense, "ACTION": "Create"})
    ET.SubElement(ET.SubElement(ledger, "NAME.LIST"), "NAME").text = suspense
    ET.SubElement(ledger, "PARENT").text = "Suspense A/c"

    df.columns = [str(c).strip().lower() for c in df.columns]
    for _, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        amount, debit, credit = 0, 0, 0
        for col in df.columns:
            val = str(row[col]).replace(',', '').strip()
            try:
                if any(x in col for x in ['debit', 'dr', 'withdrawal']): debit = float(val)
                elif any(x in col for x in ['credit', 'cr', 'deposit']): credit = float(val)
                elif 'amount' in col: amount = float(val)
            except: pass

        amount = debit if debit > 0 else credit
        if amount == 0: continue

        vch_type = "Payment" if debit > 0 else "Receipt"
        voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
        
        l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(l1, "LEDGERNAME").text = suspense
        ET.SubElement(l1, "AMOUNT").text = str(-amount if debit > 0 else amount)

        l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
        ET.SubElement(l2, "AMOUNT").text = str(amount if debit > 0 else -amount)

    return ET.tostring(envelope, encoding="utf-8", xml_declaration=True)

# --- 3. ROUTES ---
@app.route('/')
def index(): return render_template('index.html')

@app.route('/check_lock', methods=['POST'])
def check_lock():
    file = request.files['file']
    filename = secure_filename(file.filename)
    path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{filename}")
    file.save(path)
    status = "unlocked"
    try:
        if filename.lower().endswith('.pdf'):
            with pdfplumber.open(path) as pdf: _ = len(pdf.pages)
    except: status = "locked"
    finally:
        if os.path.exists(path): os.remove(path)
    return jsonify({'status': status})

@app.route('/verify_password', methods=['POST'])
def verify_password():
    file = request.files['file']
    password = request.form.get('password', '')
    path = os.path.join(app.config['UPLOAD_FOLDER'], f"v_{secure_filename(file.filename)}")
    file.save(path)
    status = "invalid"
    try:
        with pdfplumber.open(path, password=password) as pdf:
            _ = pdf.pages[0].extract_text()
        status = "valid"
    except Exception: status = "invalid"
    finally:
        if os.path.exists(path): os.remove(path)
    return jsonify({'status': status})

@app.route('/convert', methods=['POST'])
def convert():
    try:
        file = request.files['file']
        main_ledger = request.form.get('main_ledger', 'Bank Account')
        pdf_password = request.form.get('password', None)
        path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(path)
        df = parse_pdf_to_dataframe(path, pdf_password)
        xml_data = generate_tally_xml(df, main_ledger)
        return send_file(io.BytesIO(xml_data), as_attachment=True, download_name=f"{os.path.splitext(file.filename)[0]}.xml")
    except Exception as e: return jsonify({'error': str(e)}), 500

if __name__ == '__main__': app.run(debug=True)