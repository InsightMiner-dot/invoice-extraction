
import streamlit as st
import fitz  # PyMuPDF
import base64
import instructor
import os
import sqlite3
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional, Dict
import time
import io
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables from the backend
load_dotenv(override=True)

# ==============================================================
# 0. Database Setup (NEW AUDIT FEATURE)
# ==============================================================

AUDIT_FOLDER = "audit"
DB_PATH = os.path.join(AUDIT_FOLDER, "qc_master_database.sqlite")

def init_db():
    """Ensures the audit folder and SQLite database exist."""
    os.makedirs(AUDIT_FOLDER, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Create the master QC table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS qc_audit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            extraction_date TEXT,
            extraction_time TEXT,
            file_name TEXT,
            vendor_name TEXT,
            invoice_number TEXT,
            origin TEXT,
            destination TEXT,
            status TEXT,
            reason_for_review TEXT,
            extracted_total REAL,
            calculated_sum REAL,
            variance REAL
        )
    ''')
    conn.commit()
    conn.close()

def insert_audit_record(record: tuple):
    """Inserts a new row into the QC master database."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO qc_audit (
            extraction_date, extraction_time, file_name, vendor_name, 
            invoice_number, origin, destination, status, reason_for_review, 
            extracted_total, calculated_sum, variance
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', record)
    conn.commit()
    conn.close()

# Initialize the database when the app starts
init_db()


# ==============================================================
# 1. Define Data Schema (UPDATED FOR STRICT PAGING & FEES)
# ==============================================================

class LineItem(BaseModel):
    page_number: Optional[int] = Field(None, description="Page number (starting at 1)")
    material: Optional[str] = Field(None, description="Material code, part number, or SKU")
    description: str = Field(description="Name or description of the item. If a tax or fee (e.g., 'Federal Excise Tax', 'Franchise Tax') is printed as a row INSIDE the main table, extract it here as a regular line item.")
    quantity: Optional[float] = Field(None, description="Number of items SHIPPED. STRICT RULE: If you see 'QTY B/O' (Backordered) alongside 'QTY SHP', ONLY extract the Shipped amount. Do not extract backordered quantities here.")
    uom: Optional[str] = Field(None, description="Unit of Measure (e.g., EA, LBS, KG). Look for headers like 'UOM' or 'Bin'.")
    uom_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    unit_price: Optional[float] = Field(None, description="Price of a single unit")
    line_total: Optional[float] = Field(None, description="Total cost for this specific line item. Look for 'Amount', 'Total', 'Extended Price', or 'Extd Price'. STRICT RULE: If the column is blank (e.g., due to a backordered item), leave this null or 0.0. Do not guess.")

class TaxItem(BaseModel):
    tax_name: str = Field(description="The exact printed name of the tax (e.g., 'GST/HST', 'TPS/TVH', 'QST').")
    tax_amount: float = Field(description="The amount for this specific tax.")

class FeeItem(BaseModel):
    fee_name: str = Field(description="The exact printed name of the fee (e.g., 'SHOP Supplies', 'Environmental Fee', 'Pallet Charge').")
    fee_amount: float = Field(description="The amount for this specific fee.")

class InvoiceData(BaseModel):
    vendor_name: str = Field(description="Name of the company issuing the invoice")
    vendor_address: Optional[str] = Field(None, description="The FULL complete address of the vendor including street, city, state/province, and postal code. Do not extract partial addresses.")
    bill_to: Optional[str] = Field(None, description="The FULL complete 'Bill To' or 'Sold To' address including street, city, state/province, and postal code.")
    remit_to: Optional[str] = Field(None, description="The FULL complete 'Remit To' address including street, city, state/province, and postal code.")
    
    origin: Optional[str] = Field(None, description="The FULL origin physical address. STRICT GUARDRAIL: ONLY extract this if it is explicitly labeled with tags like 'Ship From', 'Origin', 'From', or 'Pickup'. If these specific tags are missing, or if it is just a random secondary address, you MUST return null. Do NOT extract short alphanumeric codes or tank numbers.")
    origin_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    
    destination: Optional[str] = Field(None, description="The FULL destination physical address ('Ship To', 'To', 'Deliver To', 'Consignee'). STRICT RULE: Do NOT extract short alphanumeric facility codes or building numbers. If a full physical address is not present, leave null.")
    destination_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    
    invoice_number: Optional[str] = Field(None, description="Unique invoice number. STRICT RULE: Only extract values explicitly labeled as 'Invoice Number' or 'Invoice #'. Do NOT extract Ticket No, Reference, Order ID, or Statement numbers.")
    invoice_number_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    
    date: Optional[str] = Field(None, description="Date the invoice was issued")
    currency: Optional[str] = Field(None, description="3-letter currency code")
    
    subtotal: Optional[float] = Field(None, description="The subtotal amount before taxes and shipping are added.")
    
    taxes: List[TaxItem] = Field(default_factory=list, description="STRICT RULE: Extract individual taxes ONLY from the summary block at the bottom of the invoice. CRITICAL GUARDRAIL: Before adding a tax here, verify you did not already extract it in `line_items`. No duplicates allowed. If a tax is already a line item, skip it here.")
    
    additional_fees: List[FeeItem] = Field(default_factory=list, description="STRICT RULE: ONLY extract fees from the summary block at the bottom. CRITICAL GUARDRAIL: Before adding a fee here, verify you did not already extract it in `line_items`. No duplicates allowed.")
    
    shipping_name: Optional[str] = Field(None, description="The exact printed name of the shipping charge (e.g., 'Freight', 'Handling', 'Delivery Fee').")
    shipping_handling: Optional[float] = Field(0.0, description="STRICT RULE: ONLY extract this if it appears in the final summary block at the bottom of the invoice, AFTER the main line items and subtotal.")
    
    total_amount: float = Field(description="Final total amount charged on the invoice")
    total_amount_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    
    custom_fields: Dict[str, Optional[str]] = Field(
        default_factory=dict, 
        description="Extract any custom fields requested by the user. The keys MUST match the requested column names exactly."
    )
    
    line_items: List[LineItem] = Field(description="List of all individual items purchased")

class InvoiceDocument(BaseModel):
    invoices: List[InvoiceData] = Field(description="List of distinct invoices. STRICT PAGING RULE: 1) If an invoice table extends across multiple pages but shares the SAME Invoice Number, MERGE all line items into ONE invoice record. 2) If the Invoice Number CHANGES on a new page, split it into a NEW separate invoice record in this list.")

# ==============================================================
# 2. PDF to Image Conversion & Extraction
# ==============================================================

def pdf_to_base64_images(file_bytes: bytes, max_pages: int = 15) -> List[str]:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    base64_images = []
    
    total_pages = len(doc)
    if total_pages > max_pages:
        st.warning(f"PDF is {total_pages} pages long. Truncating to first {max_pages} pages to prevent API timeout.")
        
    pages_to_process = min(total_pages, max_pages)
    for page_num in range(pages_to_process):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=300) 
        base64_images.append(base64.b64encode(pix.tobytes("jpeg")).decode('utf-8'))
    return base64_images

def extract_invoice_data(client, deployment: str, file_bytes: bytes, custom_cols: List[str]) -> InvoiceDocument:
    base64_images = pdf_to_base64_images(file_bytes)
    
    system_prompt = (
        "You are an expert accountant processing a document that may contain multiple distinct invoices. "
        "STRICT PAGING RULES: 1) If the same invoice number continues across multiple pages, combine all line items, taxes, and totals into a SINGLE invoice record. "
        "2) If you see a NEW invoice number, start a NEW invoice record. CRITICAL RULE AGAINST DUPLICATES: Never extract the same tax or fee twice. "
        "Your extracted total for EACH invoice must mathematically equal the calculated sum of unique items for that invoice."
    )
    
    if custom_cols:
        system_prompt += f"\n\nSTRICT RULE: The user has requested custom data extraction. You MUST also search the invoice for the following fields and place them in the 'custom_fields' dictionary: {', '.join(custom_cols)}. If a field is not found, leave its value as null."

    content_array = [{"type": "text", "text": "Extract data from this multi-page document. Pay close attention to invoice numbers. Evaluate confidence ('High', 'Medium', 'Low'). Note page numbers."}]
    for img_base64 in base64_images:
        content_array.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}})

    return client.chat.completions.create(
        model=deployment, 
        response_model=InvoiceDocument, 
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content_array}
        ]
    )

# ==============================================================
# 3. Excel Setup Utility 
# ==============================================================

def setup_excel_workbook(custom_cols: List[str]):
    wb = openpyxl.Workbook()
    
    ws_details = wb.active
    ws_details.title = "Invoice Details"
    
    details_headers = [
        "File Name", "Page #", "Vendor Name", "Vendor Address", "Bill To", "Remit To",
        "Origin", "Destination", "Invoice Number", "Date", "Currency", 
        "Material", "Description", "Quantity", "UOM", "Unit Price", "Line Total", "Subtotal", "Invoice Total",
        "Inv# Conf", "Origin Conf", "Dest Conf", "UOM Conf", "Total Conf",
        "Status", "Reason for Review"
    ]
    
    if custom_cols:
        details_headers.extend(custom_cols)
        
    ws_details.append(details_headers)
    
    ws_qc = wb.create_sheet(title="QC Summary")
    qc_headers = [
        "File Name", "Vendor Name", "Invoice Number", "Origin", "Destination", "Status", "Reason for Review", 
        "Extracted Total", "Calculated Sum (Lines+Taxes+Ship+Fees)", "Variance"
    ]
    ws_qc.append(qc_headers)
    
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for sheet in [ws_details, ws_qc]:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
    return wb, ws_details, ws_qc

# ==============================================================
# 4. Streamlit UI & Main Execution
# ==============================================================

st.set_page_config(page_title="AI Invoice Extractor", page_icon="🧾", layout="wide")

st.title("🧾 AI Invoice Extractor")
st.write("Upload PDF documents to extract structured data. Results are compiled into Excel and permanently logged to a master SQLite audit database.")

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Azure Settings")
    env_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "")
    env_key = os.getenv("AZURE_OPENAI_KEY", "")
    env_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "")
    env_api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")

    azure_endpoint = st.text_input("Endpoint URL", value=env_endpoint)
    azure_api_key = st.text_input("API Key", value=env_key, type="password")
    azure_deployment = st.text_input("Deployment Name", value=env_deployment)
    azure_api_version = st.text_input("API Version", value=env_api_version)

    st.divider()
    
    st.header("➕ Custom Extraction")
    custom_columns_input = st.text_area(
        "Custom Columns (Comma-separated)", 
        placeholder="e.g., PO Number, Cost Center, Tax ID"
    )
    custom_columns_list = [col.strip() for col in custom_columns_input.split(",")] if custom_columns_input.strip() else []

    st.divider()

    st.header("📄 Upload Files")
    uploaded_files = st.file_uploader("Upload PDF Documents", type=["pdf"], accept_multiple_files=True)

# Main Area Action
if st.button("Start Extraction", type="primary"):
    if not uploaded_files:
        st.error("Please upload at least one PDF file from the sidebar.")
    elif not all([azure_endpoint, azure_api_key, azure_deployment, azure_api_version]):
        st.error("Please ensure all Azure OpenAI settings are filled out in the sidebar.")
    else:
        try:
            client = instructor.from_openai(
                AzureOpenAI(azure_endpoint=azure_endpoint, api_key=azure_api_key, api_version=azure_api_version)
            )
        except Exception as e:
            st.error(f"Failed to initialize Azure OpenAI client: {e}")
            st.stop()

        wb, ws_details, ws_qc = setup_excel_workbook(custom_columns_list)
        red_font = Font(color="9C0006", bold=True)
        
        success_count = 0
        error_count = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_container = st.container()

        start_time = time.time()

        for idx, file in enumerate(uploaded_files):
            filename = file.name
            
            # Fetch current date and time for audit logging
            now = datetime.now()
            current_date = now.strftime("%Y-%m-%d")
            current_time = now.strftime("%H:%M:%S")

            status_text.markdown(f"**Processing Document ({idx+1}/{len(uploaded_files)}):** `{filename}`...")
            
            try:
                file_bytes = file.read()
                extracted_document = extract_invoice_data(client, azure_deployment, file_bytes, custom_columns_list)
                
                if not extracted_document.invoices:
                    log_container.warning(f"⚠️ **FLAG:** No invoices detected in `{filename}`")
                    ws_qc.append([filename, "N/A", "N/A", "N/A", "N/A", "FAIL - NO DATA", "0 Invoices Found in PDF", "", "", ""])
                    ws_qc.cell(row=ws_qc.max_row, column=6).font = red_font
                    
                    # Log failure to DB
                    insert_audit_record((
                        current_date, current_time, filename, "N/A", "N/A", "N/A", "N/A", 
                        "FAIL - NO DATA", "0 Invoices Found in PDF", 0.0, 0.0, 0.0
                    ))
                    continue
                
                for extracted_data in extracted_document.invoices:
                    calculated_line_sum = sum(item.line_total for item in extracted_data.line_items if item.line_total is not None)
                    safe_ship = extracted_data.shipping_handling if extracted_data.shipping_handling is not None else 0.0
                    safe_total_tax = sum(tax.tax_amount for tax in extracted_data.taxes if tax.tax_amount is not None)
                    safe_fees = sum(fee.fee_amount for fee in extracted_data.additional_fees if fee.fee_amount is not None)
                    
                    total_calculated = calculated_line_sum + safe_total_tax + safe_ship + safe_fees
                    variance = round(extracted_data.total_amount - total_calculated, 2)
                    
                    review_reasons = []
                    if variance != 0.0: review_reasons.append(f"Math Variance of {variance}")
                    if extracted_data.invoice_number and extracted_data.invoice_number_confidence == "Low": review_reasons.append("Low Conf: Invoice #")
                    if extracted_data.origin and extracted_data.origin_confidence == "Low": review_reasons.append("Low Conf: Origin")
                    if extracted_data.destination and extracted_data.destination_confidence == "Low": review_reasons.append("Low Conf: Destination")
                    if extracted_data.total_amount_confidence == "Low": review_reasons.append("Low Conf: Total Amount")
                    if len(extracted_data.line_items) == 0: review_reasons.append("Missing: 0 Line Items Found")
                        
                    for item in extracted_data.line_items:
                        if item.uom and item.uom_confidence == "Low":
                            short_desc = item.description[:15] + "..." if item.description and len(item.description) > 15 else item.description
                            review_reasons.append(f"Low Conf: UOM on '{short_desc}'")
                    
                    needs_review = len(review_reasons) > 0
                    status = "FAIL - NEEDS REVIEW" if needs_review else "PASS"
                    reasons_string = " | ".join(review_reasons) if needs_review else "N/A"
                    
                    def create_row(page_num, material, desc, qty, uom, uom_conf, price, line_total):
                        base_row = [
                            filename, page_num, extracted_data.vendor_name, extracted_data.vendor_address,
                            extracted_data.bill_to, extracted_data.remit_to, 
                            extracted_data.origin, extracted_data.destination, extracted_data.invoice_number,
                            extracted_data.date, extracted_data.currency, 
                            material, desc, qty, uom, price, line_total, extracted_data.subtotal, extracted_data.total_amount,
                            extracted_data.invoice_number_confidence, extracted_data.origin_confidence, 
                            extracted_data.destination_confidence, uom_conf, extracted_data.total_amount_confidence,
                            status, reasons_string
                        ]
                        custom_values = [extracted_data.custom_fields.get(col, "Not Found") for col in custom_columns_list]
                        return base_row + custom_values

                    # Append Line Items
                    if len(extracted_data.line_items) == 0:
                        ws_details.append(create_row(None, None, None, None, None, None, None, 0.0))
                    else:
                        for item in extracted_data.line_items:
                            ws_details.append(create_row(
                                item.page_number, item.material, item.description, item.quantity, 
                                item.uom, item.uom_confidence, item.unit_price, item.line_total
                            ))
                            
                    # Append Shipping, Taxes, Fees
                    if safe_ship > 0:
                        ship_label = extracted_data.shipping_name if extracted_data.shipping_name else "Shipping/Handling"
                        ws_details.append(create_row(None, None, ship_label, None, None, None, None, safe_ship))
                        
                    for tax in extracted_data.taxes:
                        if tax.tax_amount is not None and tax.tax_amount > 0:
                            ws_details.append(create_row(None, None, tax.tax_name, None, None, None, None, tax.tax_amount))
                            
                    for fee in extracted_data.additional_fees:
                        if fee.fee_amount is not None and fee.fee_amount > 0:
                            ws_details.append(create_row(None, None, fee.fee_name, None, None, None, None, fee.fee_amount))

                    # Append QC Row to Excel
                    ws_qc.append([
                        filename, extracted_data.vendor_name, extracted_data.invoice_number, 
                        extracted_data.origin, extracted_data.destination,
                        status, reasons_string, extracted_data.total_amount, total_calculated, variance
                    ])
                    
                    if needs_review:
                        ws_qc.cell(row=ws_qc.max_row, column=6).font = red_font
                        log_container.warning(f"⚠️ **FLAG:** Invoice `#{extracted_data.invoice_number}` in `{filename}` failed due to: {reasons_string}")
                    else:
                        log_container.success(f"✅ **PASS:** Invoice `#{extracted_data.invoice_number}` in `{filename}` processed perfectly.")
                        
                    # Log record to Database
                    insert_audit_record((
                        current_date, current_time, filename, 
                        extracted_data.vendor_name, extracted_data.invoice_number, 
                        extracted_data.origin, extracted_data.destination, 
                        status, reasons_string, 
                        extracted_data.total_amount, total_calculated, variance
                    ))

                    success_count += 1
                    
            except Exception as e:
                error_msg = f"System/API Crash: {str(e)}"
                log_container.error(f"❌ **CRITICAL ERROR** processing `{filename}`: {error_msg}")
                
                blank_row = [""] * 23
                error_base = [filename] + blank_row + ["FAIL - CRITICAL ERROR", error_msg]
                error_custom_padding = [""] * len(custom_columns_list)
                
                ws_details.append(error_base + error_custom_padding)
                ws_qc.append([filename, "N/A", "N/A", "N/A", "N/A", "FAIL - CRITICAL ERROR", error_msg, "", "", ""])
                ws_qc.cell(row=ws_qc.max_row, column=6).font = red_font
                
                # Log critical error to Database
                insert_audit_record((
                    current_date, current_time, filename, "N/A", "N/A", "N/A", "N/A", 
                    "FAIL - CRITICAL ERROR", error_msg, 0.0, 0.0, 0.0
                ))

                error_count += 1
            
            if idx < len(uploaded_files) - 1:
                time.sleep(3) 

            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.markdown("**Finalizing Outputs...**")

        for sheet in [ws_details, ws_qc]:
            for col in sheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                sheet.column_dimensions[column].width = min(max_length + 2, 50) 

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        end_time = time.time()
        minutes = int((end_time - start_time) // 60)
        seconds = (end_time - start_time) % 60

        status_text.empty()
        st.success(f"🎉 **Batch Processing Complete!** | Time: {minutes}m {seconds:.1f}s")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Documents Scanned", len(uploaded_files))
        col2.metric("Invoices Extracted", success_count)
        col3.metric("Critical Errors", error_count)

        st.divider()
        st.subheader("📥 Downloads")
        
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            st.download_button(
                label="📊 Download Current Batch (Excel)",
                data=excel_buffer,
                file_name=f"Extracted_Invoices_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        with col_dl2:
            # Provide download for the SQLite file
            if os.path.exists(DB_PATH):
                with open(DB_PATH, "rb") as db_file:
                    st.download_button(
                        label="🗄️ Download Master Database (SQLite)",
                        data=db_file,
                        file_name="qc_master_database.sqlite",
                        mime="application/octet-stream",
                        type="secondary"
                    )
