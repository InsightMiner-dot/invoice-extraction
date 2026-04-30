import fitz
import base64
import time
from datetime import datetime
import re
import json
from fuzzywuzzy import process
from openai import AsyncAzureOpenAI
import instructor
import os
import sqlite3
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from dotenv import load_dotenv

from core.schemas import InvoiceDocument
from core.database import insert_audit_record, load_master_suppliers, standardize_vendor, DB_PATH, AUDIT_FOLDER

load_dotenv(override=True)

client = instructor.from_openai(AsyncAzureOpenAI(
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview"),
    max_retries=3
))
DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")

async def process_single_invoice(file_bytes: bytes, filename: str, batch_id: str, max_pages: int, dpi: int, aliases: dict, custom_fields: dict):
    start_time = time.time()
    
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    base64_images = []
    total_pages = len(doc)
    pages = min(total_pages, max_pages)
    for p in range(pages):
        pix = doc[p].get_pixmap(dpi=dpi)
        base64_images.append(base64.b64encode(pix.tobytes("jpeg")).decode('utf-8'))
    doc.close()

    # Apply Settings Logic to Prompt
    sys_prompt = "Extract invoice data row by row. Start a new invoice record if invoice numbers change."
    
    if aliases:
        alias_str = "\n".join([f"- {k}: Also look for '{v}'" for k, v in aliases.items()])
        sys_prompt += f"\n\nSTANDARD FIELD ALIASES:\n{alias_str}"
        
    if custom_fields:
        rules_str = "\n".join([f"- Field Key: '{k}' | Definition & Aliases: {v}" for k, v in custom_fields.items()])
        sys_prompt += f"\n\nSTRICT CUSTOM RULES:\n{rules_str}"

    content = [{"type": "text", "text": "Extract data. Note page numbers."}]
    for img in base64_images:
        content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img}"}})

    try:
        extracted_doc = await client.chat.completions.create(
            model=DEPLOYMENT, response_model=InvoiceDocument, 
            messages=[{"role": "system", "content": sys_prompt}, {"role": "user", "content": content}]
        )
    except Exception as e:
        return {"error": str(e), "filename": filename}

    file_proc_time = round(time.time() - start_time, 2)
    master_df = load_master_suppliers()
    master_list = master_df['Original_Supplier_Name'].tolist() if not master_df.empty else []
    
    now = datetime.now()
    current_date, current_time = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")
    
    raw_pdf_text = "".join([page.get_text("text") for page in fitz.open(stream=file_bytes, filetype="pdf")])
    clean_raw_text = re.sub(r'\s+', '', raw_pdf_text).lower()

    summary_results = []
    detail_results = []

    for inv in extracted_doc.invoices:
        raw_vendor = inv.vendor_name
        if master_list and raw_vendor not in ['N/A', 'ERROR', '']:
            match = process.extractOne(raw_vendor, master_list)
            orig_supp = match[0] if match and match[1] >= 95 else standardize_vendor(raw_vendor)
        else:
            orig_supp = standardize_vendor(raw_vendor)

        calc_line_sum = sum(item.line_total for item in inv.line_items if item.line_total)
        total_calc = calc_line_sum + (inv.shipping_handling or 0.0) + sum(t.tax_amount for t in inv.taxes if t.tax_amount)
        variance = round(inv.total_amount - total_calc, 2)

        reasons = []
        if variance != 0.0: reasons.append(f"Math Variance of {variance}")
        needs_review = len(reasons) > 0
        status = "FAIL - NEEDS REVIEW" if needs_review else "PASS"
        
        # Serialize custom fields for DB storage
        custom_json = json.dumps(inv.custom_fields) if inv.custom_fields else "{}"

        # Insert to DB (34 Columns now)
        insert_audit_record((
            current_date, current_time, filename, raw_vendor, orig_supp, inv.vendor_address, inv.bill_to, inv.remit_to, inv.invoice_number, inv.date, inv.currency, 
            str(inv.origin), None, None, str(inv.destination), None, None, inv.subtotal, inv.shipping_handling or 0.0, inv.total_amount, total_calc, variance, 
            inv.invoice_number_confidence, inv.origin_confidence, inv.destination_confidence, inv.total_amount_confidence, "N/A", status, " | ".join(reasons) if needs_review else "N/A", 
            file_proc_time, total_pages, batch_id, custom_json
        ))

        summary_results.append({
            "File Name": filename, "Vendor Name": raw_vendor, "Invoice #": inv.invoice_number,
            "Variance": f"${variance:,.2f}", "Status": "✅ PASS" if status == "PASS" else "⚠️ FAIL", "Proc Time": f"{file_proc_time}s"
        })
        
        # Pass custom fields back to UI
        custom_cols = list(custom_fields.keys())
        for item in inv.line_items:
            row_data = {
                "File Name": filename, "Page #": item.page_number, "Vendor Name": raw_vendor, 
                "Original Supplier": orig_supp, "Invoice Number": inv.invoice_number, "Material": item.material,
                "Description": item.description, "Qty": item.quantity, "UOM": item.uom, "Price": item.unit_price, 
                "Line Total": item.line_total, "Origin": inv.origin, "Dest": inv.destination
            }
            # Add custom data to frontend payload
            for col in custom_cols:
                row_data[col] = inv.custom_fields.get(col, "Not Found") if inv.custom_fields else "Not Found"
            
            detail_results.append(row_data)

    return {"summary": summary_results, "details": detail_results}

def generate_excel_from_db(batch_id: str, custom_cols: list):
    os.makedirs(AUDIT_FOLDER, exist_ok=True)
    excel_path = os.path.join(AUDIT_FOLDER, f"{batch_id}.xlsx")
    
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query(f"SELECT * FROM qc_audit WHERE batch_id = '{batch_id}'", conn)
    conn.close()

    if df.empty: return False

    wb = openpyxl.Workbook()
    
    # Sheet 1: Details (Dynamic Headers)
    ws_details = wb.active
    ws_details.title = "Invoice Details"
    details_headers = ["file_name", "page_count", "vendor_name", "original_supplier_name", "invoice_number", "extracted_total", "status", "reason_for_review"]
    
    # Append Custom Columns to headers
    excel_headers = details_headers.copy()
    if custom_cols: excel_headers.extend(custom_cols)
    ws_details.append(excel_headers)
    
    for _, row in df.iterrows():
        base_row = [row[h] for h in details_headers if h in df.columns]
        
        # Parse JSON and append custom field values
        cf_data = {}
        if 'custom_fields' in df.columns and pd.notna(row['custom_fields']):
            try: cf_data = json.loads(row['custom_fields'])
            except: pass
            
        for col in custom_cols:
            base_row.append(cf_data.get(col, "Not Found"))
            
        ws_details.append(base_row)

    # Sheet 2: Summary
    ws_qc = wb.create_sheet(title="QC Summary")
    qc_headers = ["file_name", "vendor_name", "invoice_number", "origin", "destination", "status", "extracted_total", "variance"]
    ws_qc.append(qc_headers)
    for _, row in df.iterrows():
        ws_qc.append([row[h] for h in qc_headers if h in df.columns])

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for sheet in [ws_details, ws_qc]:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]: cell.fill = header_fill; cell.font = header_font

    wb.save(excel_path)
    return True
