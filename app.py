import os
import io
import zipfile
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
import xml.etree.ElementTree as ET
import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = '/tmp/uploads' 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- SMART PDF PARSER (Multiple Bank Support) ---
def parse_pdf_to_dataframe(pdf_path, pdf_password=None):
    all_data = []
    try:
        pwd = pdf_password if pdf_password else ""
        with pdfplumber.open(pdf_path, password=pwd) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
                            if any(cleaned_row): all_data.append(cleaned_row)
    except Exception: raise ValueError("PASSWORD_REQUIRED")

    if not all_data: return pd.DataFrame()

    # Smart Header Detection Logic
    header_index = 0
    max_score = 0
    keywords = ['date', 'particulars', 'description', 'narration', 'debit', 'credit', 'withdrawal', 'deposit', 'balance']
    
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

# --- XML GENERATOR ---
def generate_tally_xml(df, main_ledger_name):
    suspense = "Suspense Account"
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    # Simple Logic for Bank Entries
    df.columns = [str(c).strip().lower() for c in df.columns]
    for _, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        amount, debit, credit = 0, 0, 0
        for col in df.columns:
            val = str(row[col]).replace(',', '').strip()
            try:
                if any(x in col for x in ['debit', 'dr', 'withdrawal']): debit = float(val)
                elif any(x in col for x in ['credit', 'cr', 'deposit']): credit = float(val)
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

@app.route('/')
def index(): return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('file')
    main_ledger = request.form.get('main_ledger', 'Bank Account')
    if not files or files[0].filename == '': return jsonify({'error': "No files"}), 400

    if len(files) == 1:
        file = files[0]
        filename = secure_filename(file.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(path)
        df = parse_pdf_to_dataframe(path) if filename.endswith('.pdf') else pd.read_excel(path)
        xml_data = generate_tally_xml(df, main_ledger)
        return send_file(io.BytesIO(xml_data), as_attachment=True, download_name=f"{os.path.splitext(filename)[0]}.xml")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zf:
        for file in files:
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            try:
                df = parse_pdf_to_dataframe(path) if filename.endswith('.pdf') else pd.read_excel(path)
                xml_data = generate_tally_xml(df, main_ledger)
                zf.writestr(f"{os.path.splitext(filename)[0]}.xml", xml_data)
            except: continue
    zip_buffer.seek(0)
    return send_file(zip_buffer, as_attachment=True, download_name="BankFlow_Bulk.zip")

if __name__ == '__main__': app.run(debug=True)