import fitz
import base64
import asyncio
from datetime import datetime
import re
from fuzzywuzzy import process
from openai import AsyncAzureOpenAI
import instructor
import os
from dotenv import load_dotenv

from core.schemas import InvoiceDocument
from core.database import insert_audit_record, load_master_suppliers, standardize_vendor

load_dotenv(override=True)

client = instructor.from_openai(AsyncAzureOpenAI(
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview"),
    max_retries=3
))
DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")

def get_raw_pdf_text(file_bytes: bytes) -> str:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = "".join([page.get_text("text") + " " for page in doc])
    doc.close()
    return text

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
    tasks = [extract_single_invoice(fb, max_pages, dpi) for fb in file_bytes_list]
    results = await asyncio.gather(*tasks, return_exceptions=True)
    
    master_df = load_master_suppliers()
    master_list = master_df['Original_Supplier_Name'].tolist() if not master_df.empty else []
    batch_id = datetime.now().strftime("BATCH_%Y%m%d_%H%M%S")
    now = datetime.now()
    current_date, current_time = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")

    processed_data = []

    for idx, result in enumerate(results):
        filename = file_names[idx]
        if isinstance(result, Exception):
            insert_audit_record((
                current_date, current_time, filename, "ERROR", "ERROR", "ERROR", "ERROR", "ERROR", 
                "ERROR", "ERROR", "ERROR", "ERROR", None, None, "ERROR", None, None, 
                0.0, 0.0, 0.0, 0.0, 0.0, "ERROR", "ERROR", "ERROR", "ERROR", "ERROR", 
                "FAIL - CRITICAL ERROR", str(result), 0.0, 0, batch_id
            ))
            continue
            
        extracted_doc, total_pages = result
        raw_pdf_text = get_raw_pdf_text(file_bytes_list[idx])
        clean_raw_text = re.sub(r'\s+', '', raw_pdf_text).lower()

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
            if inv.invoice_number and re.sub(r'\s+', '', inv.invoice_number).lower() not in clean_raw_text:
                reasons.append("HALLUCINATION: Inv # not in doc")
            
            needs_review = len(reasons) > 0
            status = "FAIL - NEEDS REVIEW" if needs_review else "PASS"

            insert_audit_record((
                current_date, current_time, filename, raw_vendor, orig_supp, 
                inv.vendor_address, inv.bill_to, inv.remit_to, inv.invoice_number, inv.date, inv.currency, 
                str(inv.origin), None, None, str(inv.destination), None, None, 
                inv.subtotal, inv.shipping_handling or 0.0, inv.total_amount, total_calc, variance, 
                inv.invoice_number_confidence, inv.origin_confidence, inv.destination_confidence, 
                inv.total_amount_confidence, "N/A", status, " | ".join(reasons) if needs_review else "N/A", 
                0.0, total_pages, batch_id
            ))
            processed_data.append({"file": filename, "vendor": orig_supp, "status": status, "variance": variance})

    return processed_data
