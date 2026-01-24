import xml.etree.ElementTree as ET

def generate_tally_xml(transactions):
    # Root Envelope
    envelope = ET.Element("ENVELOPE")
    
    # Header
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    # Body
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(req_desc, "REPORTNAME").text = "Vouchers"
    
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    for trans in transactions:
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE")
        # Voucher Tag
        voucher = ET.SubElement(tally_msg, "VOUCHER", VCHTYPE=trans['voucher_type'], ACTION="Create")
        
        # Date
        ET.SubElement(voucher, "DATE").text = trans['date']
        ET.SubElement(voucher, "NARRATION").text = trans['narration']
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = trans['voucher_type']
        
        # Logic for Double Entry
        # Entry 1: Bank Ledger (Credit in Payment, Debit in Receipt)
        bank_entry = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(bank_entry, "LEDGERNAME").text = "Bank Account" # User ka Bank Ledger Name
        
        if trans['voucher_type'] == "Payment":
            ET.SubElement(bank_entry, "ISDEEMEDPOSITIVE").text = "No" # Credit
            ET.SubElement(bank_entry, "AMOUNT").text = str(trans['amount'])
        else:
            ET.SubElement(bank_entry, "ISDEEMEDPOSITIVE").text = "Yes" # Debit
            ET.SubElement(bank_entry, "AMOUNT").text = str(-trans['amount']) # Negative for Debit in Tally XML logic sometimes varies, but standard is -ve for debit in some schemas. Let's keep absolute for simple import.
            # Correction: Tally XML usually takes negative amount for Debit and positive for Credit in older versions, but let's stick to ISDEEMEDPOSITIVE flag which is safer.
            
            # Re-correcting Amount Logic for Tally XML standard:
            # Usually amounts are just negative numbers for credit entries if not using ISDEEMEDPOSITIVE explicitly correctly.
            # Let's use simple logic:
            # Payment: Bank (Cr) -> Amount, Expense (Dr) -> -Amount
        
        # Simplified Logic for XML:
        # Just putting amount. Tally imports based on IsDeemedPositive.
        
        # Entry 2: Suspense Ledger (Opposite of Bank)
        suspense_entry = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(suspense_entry, "LEDGERNAME").text = "Suspense A/c" # Default ledger
        
        if trans['voucher_type'] == "Payment":
            ET.SubElement(suspense_entry, "ISDEEMEDPOSITIVE").text = "Yes" # Debit
            ET.SubElement(suspense_entry, "AMOUNT").text = str(-trans['amount'])
        else:
            ET.SubElement(suspense_entry, "ISDEEMEDPOSITIVE").text = "No" # Credit
            ET.SubElement(suspense_entry, "AMOUNT").text = str(trans['amount'])

    # Convert to String
    xml_str = ET.tostring(envelope, encoding='unicode')
    return xml_str