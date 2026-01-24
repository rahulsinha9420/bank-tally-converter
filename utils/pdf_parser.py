import pdfplumber
import re

def parse_pdf_statement(filepath):
    transactions = []
    
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            
            if table:
                for row in table:
                    clean_row = [str(cell) if cell else '' for cell in row]
                    
                    try:
                        # Row length check
                        if len(clean_row) < 3: 
                            continue
                            
                        # Data extraction logic (adjust indices based on your bank PDF)
                        date_str = clean_row[0] 
                        narration = clean_row[1]
                        
                        # Basic date validation
                        if not re.search(r'\d', date_str):
                            continue
                        
                        # Amount search (simple logic for MVP)
                        # Hum last ke columns check karenge amount ke liye
                        amount = 0
                        voucher_type = ""
                        
                        # Example: Last column is Credit, 2nd Last is Debit
                        credit_str = clean_row[-1].replace(',', '')
                        debit_str = clean_row[-2].replace(',', '')

                        if debit_str and debit_str.replace('.', '').isdigit():
                            val = float(debit_str)
                            if val > 0:
                                amount = val
                                voucher_type = "Payment"
                        
                        elif credit_str and credit_str.replace('.', '').isdigit():
                            val = float(credit_str)
                            if val > 0:
                                amount = val
                                voucher_type = "Receipt"
                        
                        if amount > 0:
                            transactions.append({
                                "date": date_str.replace('/', ''),
                                "narration": narration,
                                "amount": amount,
                                "voucher_type": voucher_type
                            })
                            
                    except Exception:
                        continue
                        
    return transactions