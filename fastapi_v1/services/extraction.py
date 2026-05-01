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
    raw_pdf_text_list = []
    
    total_pages = len(doc)
    pages = min(total_pages, max_pages)
    
    for p in range(pages):
        page = doc[p]
        pix = page.get_pixmap(dpi=dpi)
        base64_images.append(base64.b64encode(pix.tobytes("jpeg")).decode('utf-8'))
        raw_pdf_text_list.append(page.get_text("text"))
        
    doc.close()
    
    clean_raw_text = re.sub(r'\s+', '', "".join(raw_pdf_text_list)).lower()

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

    summary_results = []
    detail_results = []

    for inv in extracted_doc.invoices:
        raw_vendor = inv.vendor_name
        if master_list and raw_vendor not in ['N/A', 'ERROR', '']:
            match = process.extractOne(raw_vendor, master_list)
            orig_supp = match[0] if match and match[1] >= 95 else standardize_vendor(raw_vendor)
        else:
            orig_supp = standardize_vendor(raw_vendor)

        # --- UPDATED VARIANCE MATH ---
        # Variance = Invoice Total - Sum of Line Totals
        calc_line_sum = sum(item.line_total for item in inv.line_items if item.line_total is not None)
        variance = round(inv.total_amount - calc_line_sum, 2)

        reasons = []
        if variance != 0.0: reasons.append(f"Math Variance of {variance}")
        if inv.invoice_number and re.sub(r'\s+', '', inv.invoice_number).lower() not in clean_raw_text:
            reasons.append("HALLUCINATION: Inv # not in doc")
        
        needs_review = len(reasons) > 0
        
        # --- UPDATED STATUS STRING ---
        status = "NEEDS REVIEW" if needs_review else "PASS"
        
        custom_json = json.dumps(inv.custom_fields) if inv.custom_fields else "{}"
        
        line_items_list = []
        for item in inv.line_items:
            line_items_list.append({
                "page_number": item.page_number, "material": item.material, "description": item.description,
                "quantity": item.quantity, "uom": item.uom, "uom_confidence": item.uom_confidence,
                "unit_price": item.unit_price, "line_total": item.line_total
            })
        line_items_json = json.dumps(line_items_list)

        insert_audit_record((
            current_date, current_time, filename, raw_vendor, orig_supp, inv.vendor_address, inv.bill_to, inv.remit_to, inv.invoice_number, inv.date, inv.currency, 
            str(inv.origin), None, None, str(inv.destination), None, None, inv.subtotal, inv.shipping_handling or 0.0, inv.total_amount, calc_line_sum, variance, 
            inv.invoice_number_confidence, inv.origin_confidence, inv.destination_confidence, inv.total_amount_confidence, "N/A", status, " | ".join(reasons) if needs_review else "N/A", 
            file_proc_time, total_pages, batch_id, custom_json, line_items_json
        ))

        summary_results.append({
            "File Name": filename, "Vendor Name": raw_vendor, "Invoice #": inv.invoice_number,
            "Variance": f"${variance:,.2f}", "Status": "✅ PASS" if status == "PASS" else "⚠️ NEEDS REVIEW", 
            "Proc Time": f"{file_proc_time}s", "Total Pages": total_pages
        })
        
        custom_cols = list(custom_fields.keys())
        for item in inv.line_items:
            row_data = {
                "File Name": filename, "Page #": item.page_number, "Vendor Name": raw_vendor, 
                "Original Supplier": orig_supp, "Invoice Number": inv.invoice_number, "Material": item.material,
                "Description": item.description, "Qty": item.quantity, "UOM": item.uom, "Price": item.unit_price, 
                "Line Total": item.line_total, "Origin": inv.origin, "Dest": inv.destination,
                "Inv# Conf": inv.invoice_number_confidence, "Total Conf": inv.total_amount_confidence, 
                "Proc Time": f"{file_proc_time}s", "Status": "✅ PASS" if status == "PASS" else "⚠️ NEEDS REVIEW", "Variance": variance
            }
            for col in custom_cols:
                row_data[col] = inv.custom_fields.get(col, "Not Found") if inv.custom_fields else "Not Found"
            detail_results.append(row_data)

    return {"summary": summary_results, "details": detail_results, "total_file_pages": total_pages}


def generate_excel_from_db(batch_id: str, custom_cols: list):
    os.makedirs(AUDIT_FOLDER, exist_ok=True)
    excel_path = os.path.join(AUDIT_FOLDER, f"{batch_id}.xlsx")
    
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query(f"SELECT * FROM qc_audit WHERE batch_id = '{batch_id}'", conn)
    conn.close()

    if df.empty: return False

    wb = openpyxl.Workbook()
    
    # SHEET 1: LINE ITEMS
    ws_details = wb.active
    ws_details.title = "Invoice Details"
    
    details_headers = [
        "File Name", "Page #", "Vendor Name", "Original Supplier Name", "Invoice Number", "Date", "Currency",
        "Material", "Description", "Quantity", "UOM", "Unit Price", "Line Total",
        "Subtotal", "Invoice Total", "Variance", "Inv# Conf", "Total Conf", "UOM Conf",
        "Proc Time", "Status", "Reason for Review"
    ]
    excel_headers = details_headers.copy()
    if custom_cols: excel_headers.extend(custom_cols)
    ws_details.append(excel_headers)
    
    for _, row in df.iterrows():
        cf_data = {}
        if 'custom_fields' in df.columns and pd.notna(row['custom_fields']):
            try: cf_data = json.loads(row['custom_fields'])
            except: pass
            
        if 'line_items' in df.columns and pd.notna(row['line_items']):
            try:
                line_items = json.loads(row['line_items'])
                if not line_items: line_items = [{}] 
            except: line_items = [{}]
        else:
            line_items = [{}]

        for item in line_items:
            base_row = [
                row['file_name'], item.get('page_number', ''), row['vendor_name'], row['original_supplier_name'],
                row['invoice_number'], row['invoice_date'], row['currency'],
                item.get('material', ''), item.get('description', 'NO ITEMS FOUND'), item.get('quantity', ''),
                item.get('uom', ''), item.get('unit_price', ''), item.get('line_total', ''),
                row['subtotal'], row['extracted_total'], row['variance'], row['invoice_number_conf'], 
                row['total_amount_conf'], item.get('uom_confidence', ''),
                row['processing_time'], row['status'], row['reason_for_review']
            ]
            for col in custom_cols: base_row.append(cf_data.get(col, "Not Found"))
            ws_details.append(base_row)

    # SHEET 2: QC SUMMARY (Variance is explicitly present here)
    ws_qc = wb.create_sheet(title="QC Summary")
    qc_headers = ["file_name", "vendor_name", "invoice_number", "origin", "destination", "status", "extracted_total", "variance", "processing_time", "reason_for_review"]
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
