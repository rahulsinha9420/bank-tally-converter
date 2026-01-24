import pandas as pd

def normalize_headers(df):
    """
    Ye function alag-alag bank headers ko standard format mein badal deta hai.
    Example: 'Txn Date' -> 'Date', 'Withdrawal' -> 'Debit'
    """
    # Column names ko clean karo (spaces aur case hatana)
    df.columns = df.columns.str.strip().str.lower()
    
    # Mapping Dictionary (Left side: Humara Standard Name, Right side: Bank ke variations)
    mappings = {
        'date': ['date', 'txn date', 'transaction date', 'value date'],
        'narration': ['narration', 'description', 'particulars', 'remarks', 'details'],
        'debit': ['debit', 'withdrawal', 'dr', 'debit amount', 'withdrawal amount'],
        'credit': ['credit', 'deposit', 'cr', 'credit amount', 'deposit amount']
    }
    
    # Columns rename karna
    new_columns = {}
    found_cols = df.columns.tolist()
    
    for standard_name, variations in mappings.items():
        for col in found_cols:
            if col in variations:
                new_columns[col] = standard_name.capitalize() # 'date' -> 'Date'
                break # Ek mil gaya toh break karo
                
    # DataFrame ke columns rename karo
    df.rename(columns=new_columns, inplace=True)
    return df

def parse_excel_statement(filepath):
    # Excel padhna
    df = pd.read_excel(filepath, engine='openpyxl')
    
    # 1. Headers thik karna (Smart Logic)
    df = normalize_headers(df)
    
    # 2. Check karna ki required columns mile ya nahi
    required_cols = ['Date', 'Narration']
    # Debit/Credit mein se koi ek bhi ho to chalega (kabhi kabhi 'Amount' + 'Type' hota hai, par abhi simple rakhte hain)
    
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        raise ValueError(f"Columns match nahi huye. Python dhoondh raha tha: Date, Narration. \nApke Excel mein headers hain: {list(df.columns)}")
    
    transactions = []
    df = df.fillna(0)
    
    for index, row in df.iterrows():
        try:
            # Date Handling
            raw_date = row.get('Date')
            if pd.isna(raw_date) or str(raw_date).strip() == "" or raw_date == 0:
                continue
                
            date_obj = pd.to_datetime(raw_date, errors='coerce')
            if pd.isna(date_obj):
                continue
                
            formatted_date = date_obj.strftime('%Y%m%d')
            narration = str(row.get('Narration', ''))
            
            # Amount Handling (Debit/Credit columns dhoondhna)
            debit = float(row.get('Debit', 0)) if 'Debit' in df.columns else 0
            credit = float(row.get('Credit', 0)) if 'Credit' in df.columns else 0
            
            amount = 0
            voucher_type = ""
            
            if debit > 0:
                amount = debit
                voucher_type = "Payment"
            elif credit > 0:
                amount = credit
                voucher_type = "Receipt"
            
            if amount > 0:
                transactions.append({
                    "date": formatted_date,
                    "narration": narration,
                    "amount": amount,
                    "voucher_type": voucher_type
                })
        except Exception as e:
            continue
            
    return transactions