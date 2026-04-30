import fitz
import base64
import asyncio
import time
from datetime import datetime
import re
from fuzzywuzzy import process
from openai import AsyncAzureOpenAI
import instructor
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from dotenv import load_dotenv

from core.schemas import InvoiceDocument
from core.database import insert_audit_record, load_master_suppliers, standardize_vendor, AUDIT_FOLDER

load_dotenv(override=True)

client = instructor.from_openai(AsyncAzureOpenAI(
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview"),
    max_retries=3
))
DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")

def setup_excel_workbook():
    wb = openpyxl.Workbook()
    ws_details = wb.active
    ws_details.title = "Invoice Details"
    details_headers = [
        "File Name", "Page #", "Vendor Name", "Original Supplier Name", "Vendor Address", "Bill To", "Remit To",
        "Origin", "Destination", "Invoice Number", "Date", "Currency", 
        "Material", "Description", "Quantity", "UOM", "Unit Price", "Line Total", "Subtotal", "Invoice Total",
        "Inv# Conf", "Origin Conf", "Dest Conf", "UOM Conf", "Total Conf", "Status", "Reason for Review"
    ]
    ws_details.append(details_headers)
    
    ws_qc = wb.create_sheet(title="QC Summary")
    qc_headers = [
        "File Name", "Vendor Name", "Original Supplier Name", "Invoice Number", "Origin", "Destination", "Status", "Reason for Review", 
        "Extracted Total", "Calculated Sum", "Variance"
    ]
    ws_qc.append(qc_headers)
    
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for sheet in [ws_details, ws_qc]:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal="center", vertical="center")
    return wb, ws_details, ws_qc

async def extract_single_invoice(file_bytes, max_pages, dpi):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    base64_images = []
    total_pages = len(doc)
    pages = min(total_pages, max_pages)
    for p in range(pages):
        pix = doc[p].get_pixmap(dpi=dpi)
        base64_images.append(base64.b64encode(pix.tobytes("jpeg")).decode('utf-8'))
    doc.close()

    sys_prompt = "Extract invoice data row by row. Start a new invoice record if invoice numbers change."
    content = [{"type": "text", "text": "Extract data. Note page numbers."}]
    for img in base64_images:
        content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img}"}})

    response = await client.chat.completions.create(
        model=DEPLOYMENT, response_model=InvoiceDocument, 
        messages=[{"role": "system", "content": sys_prompt}, {"role": "user", "content": content}]
    )
    return response, total_pages

async def process_batch_concurrently(file_bytes_list, file_names, max_pages=15, dpi=300):
    start_time_batch = time.time()
    tasks = [extract_single_invoice(fb, max_pages, dpi) for fb in file_bytes_list]
    results = await asyncio.gather(*tasks, return_exceptions=True)
    
    master_df = load_master_suppliers()
    master_list = master_df['Original_Supplier_Name'].tolist() if not master_df.empty else []
    
    batch_id = datetime.now().strftime("BATCH_%Y%m%d_%H%M%S")
    now = datetime.now()
    current_date, current_time = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")

    wb, ws_details, ws_qc = setup_excel_workbook()
    
    current_run_summary = []
    current_run_details = []

    for idx, result in enumerate(results):
        filename = file_names[idx]
        file_proc_time = round((time.time() - start_time_batch) / len(file_names), 2)
        
        if isinstance(result, Exception):
            current_run_summary.append({"File Name": filename, "Vendor Name": "ERROR", "Invoice #": "ERROR", "Status": "FAIL - CRITICAL ERROR", "Reason": str(result)})
            continue
            
        extracted_doc, total_pages = result

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
            reasons_str = " | ".join(reasons) if needs_review else "N/A"

            # 1. DB Insert
            insert_audit_record((
                current_date, current_time, filename, raw_vendor, orig_supp, inv.vendor_address, inv.bill_to, inv.remit_to, inv.invoice_number, inv.date, inv.currency, 
                str(inv.origin), None, None, str(inv.destination), None, None, inv.subtotal, inv.shipping_handling or 0.0, inv.total_amount, total_calc, variance, 
                inv.invoice_number_confidence, inv.origin_confidence, inv.destination_confidence, inv.total_amount_confidence, "N/A", status, reasons_str, file_proc_time, total_pages, batch_id
            ))

            # 2. Append to Details Excel Sheet & Array
            def create_row_dict(page_num, material, desc, qty, uom, uom_conf, price, line_total):
                return {
                    "File Name": filename, "Page #": page_num, "Vendor Name": raw_vendor, "Original Supplier Name": orig_supp,
                    "Vendor Address": inv.vendor_address, "Bill To": inv.bill_to, "Remit To": inv.remit_to,
                    "Origin": inv.origin, "Destination": inv.destination, "Invoice Number": inv.invoice_number, 
                    "Date": inv.date, "Currency": inv.currency, "Material": material, "Description": desc, 
                    "Quantity": qty, "UOM": uom, "Unit Price": price, "Line Total": line_total,
                    "Subtotal": inv.subtotal, "Invoice Total": inv.total_amount,
                    "Inv# Conf": inv.invoice_number_confidence, "Origin Conf": inv.origin_confidence,
                    "Dest Conf": inv.destination_confidence, "UOM Conf": uom_conf, "Total Conf": inv.total_amount_confidence,
                    "Status": status, "Reason for Review": reasons_str
                }

            if not inv.line_items:
                row_dict = create_row_dict(None, None, "NO ITEMS FOUND", None, None, None, None, 0.0)
                ws_details.append(list(row_dict.values()))
                current_run_details.append(row_dict)
            else:
                for item in inv.line_items:
                    row_dict = create_row_dict(item.page_number, item.material, item.description, item.quantity, item.uom, item.uom_confidence, item.unit_price, item.line_total)
                    ws_details.append(list(row_dict.values()))
                    current_run_details.append(row_dict)

            # 3. Append to QC Summary Excel Sheet & Array
            ws_qc.append([filename, raw_vendor, orig_supp, inv.invoice_number, inv.origin, inv.destination, status, reasons_str, inv.total_amount, total_calc, variance])
            current_run_summary.append({
                "File Name": filename, "Vendor Name": raw_vendor, "Invoice #": inv.invoice_number,
                "Origin": inv.origin or "Missing", "Destination": inv.destination or "Missing",
                "Variance": f"${variance:,.2f}", "Status": "✅ PASS" if status == "PASS" else "⚠️ FAIL",
                "Proc Time": f"{file_proc_time}s"
            })

    # Save physical file to disk for downloading
    os.makedirs(AUDIT_FOLDER, exist_ok=True)
    excel_path = os.path.join(AUDIT_FOLDER, f"{batch_id}.xlsx")
    wb.save(excel_path)

    return {
        "batch_id": batch_id,
        "summary": current_run_summary,
        "details": current_run_details
    }
