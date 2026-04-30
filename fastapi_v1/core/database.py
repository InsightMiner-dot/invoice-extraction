import sqlite3
import os
import pandas as pd
import re

AUDIT_FOLDER = "audit"
DB_PATH = os.path.join(AUDIT_FOLDER, "qc_master_database.sqlite")
MASTER_CSV_PATH = os.path.join(AUDIT_FOLDER, "master_suppliers.csv")

def init_db():
    os.makedirs(AUDIT_FOLDER, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS qc_audit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            extraction_date TEXT, extraction_time TEXT, file_name TEXT,
            vendor_name TEXT, original_supplier_name TEXT, vendor_address TEXT,
            bill_to TEXT, remit_to TEXT, invoice_number TEXT, invoice_date TEXT,
            currency TEXT, origin TEXT, suggested_origin TEXT, final_origin TEXT,
            destination TEXT, suggested_destination TEXT, final_destination TEXT,
            subtotal REAL, shipping_handling REAL, extracted_total REAL,
            calculated_sum REAL, variance REAL, invoice_number_conf TEXT,
            origin_conf TEXT, destination_conf TEXT, total_amount_conf TEXT,
            uom_conf TEXT, status TEXT, reason_for_review TEXT,
            processing_time REAL, page_count INTEGER, batch_id TEXT,
            custom_fields TEXT, line_items TEXT
        )
    ''')
    
    # Safe migration: Add columns if they are missing from an older database
    try: cursor.execute("ALTER TABLE qc_audit ADD COLUMN custom_fields TEXT")
    except: pass
    try: cursor.execute("ALTER TABLE qc_audit ADD COLUMN line_items TEXT")
    except: pass
        
    conn.commit()
    conn.close()

def fetch_audit_data() -> pd.DataFrame:
    if not os.path.exists(DB_PATH): return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM qc_audit", conn)
    conn.close()
    return df

def insert_audit_record(record: tuple):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # Now inserting 35 values
    cursor.execute('''
        INSERT INTO qc_audit (
            extraction_date, extraction_time, file_name, vendor_name, original_supplier_name, 
            vendor_address, bill_to, remit_to, invoice_number, invoice_date, currency, 
            origin, suggested_origin, final_origin, destination, suggested_destination, final_destination, 
            subtotal, shipping_handling, extracted_total, calculated_sum, variance, 
            invoice_number_conf, origin_conf, destination_conf, total_amount_conf, uom_conf, 
            status, reason_for_review, processing_time, page_count, batch_id, custom_fields, line_items
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', record)
    conn.commit()
    conn.close()

def load_master_suppliers() -> pd.DataFrame:
    if os.path.exists(MASTER_CSV_PATH):
        try: return pd.read_csv(MASTER_CSV_PATH, encoding='utf-8-sig')[['Original_Supplier_Name']].dropna()
        except: pass
    return pd.DataFrame(columns=["Original_Supplier_Name"])

def standardize_vendor(name):
    if not isinstance(name, str) or name in ['N/A', 'ERROR', '']: return name
    name = name.upper()
    name = re.sub(r'[.,]', '', name)
    return re.sub(r'\b(LLC|INC|LTD|CORP|CORPORATION|CO|COMPANY)\b', '', name).strip()
