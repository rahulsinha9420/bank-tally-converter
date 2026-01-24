import os
import pandas as pd
from flask import Flask, render_template, request, send_file
import xml.etree.ElementTree as ET
import pdfplumber
from werkzeug.utils import secure_filename

app = Flask(__name__)

# --- CONFIGURATION ---
# '/tmp' folder Render server par sabse safe hota hai temporary files ke liye
UPLOAD_FOLDER = '/tmp/uploads' 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- 1. PDF PARSER FUNCTION ---
def parse_pdf_to_dataframe(pdf_path):
    all_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            # Data saaf karna (None ko hatana)
                            cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
                            # Agar row khali nahi hai to add karo
                            if any(cleaned_row): 
                                all_data.append(cleaned_row)
    except Exception as e:
        print(f"PDF Parsing Error: {e}")
        return pd.DataFrame()

    if not all_data: return pd.DataFrame()

    headers = all_data[0]
    data = all_data[1:]
    
    # Headers ko unique banana (taaki error na aaye)
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
    
    # Column names ko clean karo
    df.columns = [str(c).strip() for c in df.columns]
    
    for index, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        
        # --- Date Logic ---
        date_val = ""
        for col in df.columns:
            if "date" in col.lower():
                date_val = row[col]
                break
        try:
            date_obj = pd.to_datetime(date_val, dayfirst=True)
            tally_date = date_obj.strftime('%Y%m%d')
        except:
            tally_date = "20240401" # Default agar date fail ho jaye

        # --- Amount Logic ---
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
        
        # --- Narration Logic ---
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

        # --- XML Structure ---
        if conversion_type == 'sales':
            vch_type = "Sales"
            is_deemed_positive_party = "Yes"
            is_deemed_positive_main = "No"
        elif conversion_type == 'purchase':
            vch_type = "Purchase"
            is_deemed_positive_party = "No"
            is_deemed_positive_main = "Yes"
        else: # Bank Statement
            if debit > 0:
                vch_type = "Payment"
                is_party_debit = "Yes"
                is_bank_credit = "No"
            else:
                vch_type = "Receipt"
                is_party_debit = "No"
                is_bank_credit = "Yes"
            
            # Bank specific structure
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

        # Sales/Purchase Structure
        voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
        ET.SubElement(voucher, "DATE").text = tally_date
        ET.SubElement(voucher, "NARRATION").text = narration
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
        
        l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(l1, "LEDGERNAME").text = party_name
        ET.SubElement(l1, "ISDEEMEDPOSITIVE").text = is_deemed_positive_party
        ET.SubElement(l1, "AMOUNT").text = str(-amount if is_deemed_positive_party == "Yes" else amount)

        l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
        ET.SubElement(l2, "ISDEEMEDPOSITIVE").text = is_deemed_positive_main
        ET.SubElement(l2, "AMOUNT").text = str(-amount if is_deemed_positive_main == "Yes" else amount)

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
            # Secure Filename (Safety ke liye)
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], "Tally_Import.xml")
            
            file.save(input_path)
            
            df = pd.DataFrame()
            fname_lower = filename.lower()

            # Format Check
            if fname_lower.endswith('.pdf'):
                df = parse_pdf_to_dataframe(input_path)
                if df.empty: 
                    return "<h1>Error:</h1><p>PDF se data nahi nikala ja saka. Kripya ise Excel mein convert karke upload karein.</p>"
            
            elif fname_lower.endswith(('.xls', '.xlsx')):
                engine = 'xlrd' if fname_lower.endswith('.xls') else 'openpyxl'
                df = pd.read_excel(input_path, engine=engine)
            else:
                return "<h1>Format Error:</h1><p>Sirf PDF ya Excel (.xls, .xlsx) files allowed hain.</p>"

            df = df.fillna('')
            generate_tally_xml(df, output_path, conversion_type, main_ledger, suspense_ledger)
            
            return send_file(output_path, as_attachment=True)
            
    except Exception as e:
        # Asli Error dikhayega
        return f"<h1>Something went wrong!</h1><p>Error Details: {str(e)}</p>"

if __name__ == '__main__':
    app.run(debug=True)