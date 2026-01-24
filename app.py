import os
import pandas as pd
from flask import Flask, render_template, request, send_file
import xml.etree.ElementTree as ET

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def generate_tally_xml(df, output_path, conversion_type, main_ledger_name, suspense_ledger_name):
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(req_desc, "REPORTNAME").text = "Vouchers"
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    for index, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        
        # --- Common Data Cleaning ---
        try:
            date_obj = pd.to_datetime(row.get('Date'), dayfirst=True)
            tally_date = date_obj.strftime('%Y%m%d')
        except:
            tally_date = "20240401"
            
        try:
            if 'Amount' in df.columns:
                amount = float(str(row['Amount']).replace(',', ''))
            else:
                debit = float(str(row.get('Debit', 0)).replace(',', ''))
                credit = float(str(row.get('Credit', 0)).replace(',', ''))
                amount = debit if debit > 0 else credit
        except:
            amount = 0

        narration = str(row.get('Narration', ''))
        party_name_from_excel = str(row.get('Party', str(row.get('Particulars', ''))))

        if conversion_type == 'sales':
            vch_type = "Sales"
            voucher = ET.SubElement(tally_msg, "VOUCHER", {"VCHTYPE": vch_type, "ACTION": "Create"})
            ET.SubElement(voucher, "DATE").text = tally_date
            ET.SubElement(voucher, "NARRATION").text = narration
            ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type

            l1 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l1, "LEDGERNAME").text = party_name_from_excel
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
            ET.SubElement(l1, "LEDGERNAME").text = party_name_from_excel
            ET.SubElement(l1, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(l1, "AMOUNT").text = str(amount) 

            l2 = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l2, "LEDGERNAME").text = main_ledger_name
            ET.SubElement(l2, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(l2, "AMOUNT").text = str(-amount)

        else:
            debit_amt = float(str(row.get('Debit', 0)).replace(',', ''))
            if debit_amt > 0:
                vch_type = "Payment"
                is_party_debit = "Yes"
                is_bank_credit = "No"
                final_amt = debit_amt
            else:
                vch_type = "Receipt"
                is_party_debit = "No"
                is_bank_credit = "Yes"
                final_amt = float(str(row.get('Credit', 0)).replace(',', ''))

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
            # --- YAHAN CHANGE KIYA HAI ---
            # Engine 'openpyxl' bata diya taaki error na aaye
            df = pd.read_excel(input_path, engine='openpyxl') 
            df = df.fillna('')
            
            generate_tally_xml(df, output_path, conversion_type, main_ledger, suspense_ledger)
            
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"Error: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)