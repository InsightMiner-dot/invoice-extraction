import fitz  # PyMuPDF
import base64
import instructor
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional

# ==============================================================
# 1. Define Data Schema 
# ==============================================================

class LineItem(BaseModel):
    page_number: Optional[int] = Field(None, description="Page number (starting at 1)")
    material: Optional[str] = Field(None, description="Material code, part number, or SKU")
    description: str = Field(description="Name or description of the item")
    quantity: Optional[float] = Field(None, description="Number of items purchased")
    uom: Optional[str] = Field(None, description="Unit of Measure")
    uom_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    unit_price: Optional[float] = Field(None, description="Price of a single unit")
    line_total: float = Field(description="Total cost for this specific line item.")

class InvoiceData(BaseModel):
    vendor_name: str = Field(description="Name of the company issuing the invoice")
    vendor_address: Optional[str] = Field(None, description="Full address of the vendor")
    bill_to: Optional[str] = Field(None, description="'Bill To' or 'Sold To' address")
    ship_to: Optional[str] = Field(None, description="'Ship To' address.")
    remit_to: Optional[str] = Field(None, description="'Remit To' address")
    origin: Optional[str] = Field(None, description="Origin address ('From', 'Ship From', 'Pickup', 'Generator').")
    origin_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    destination: Optional[str] = Field(None, description="Destination address ('To', 'Deliver To', 'Consignee').")
    destination_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    invoice_number: Optional[str] = Field(None, description="Unique invoice ID or number")
    invoice_number_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    date: Optional[str] = Field(None, description="Date the invoice was issued")
    currency: Optional[str] = Field(None, description="3-letter currency code")
    tax_amount: Optional[float] = Field(0.0, description="Total tax amount. Default to 0.0 if not found.")
    shipping_handling: Optional[float] = Field(0.0, description="Shipping/Freight charges. Default to 0.0 if not found.")
    total_amount: float = Field(description="Final total amount charged on the invoice")
    total_amount_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    line_items: List[LineItem] = Field(description="List of all individual items purchased")

# ==============================================================
# 2. Setup Azure OpenAI Client
# ==============================================================

azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
azure_api_key = os.getenv("AZURE_OPENAI_KEY")
azure_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT")
azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION")

client = instructor.from_openai(
    AzureOpenAI(azure_endpoint=azure_endpoint, api_key=azure_api_key, api_version=azure_api_version)
)

# ==============================================================
# 3. PDF to Image Conversion & Extraction
# ==============================================================

def pdf_to_base64_images(pdf_path: str) -> List[str]:
    doc = fitz.open(pdf_path)
    base64_images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        base64_images.append(base64.b64encode(pix.tobytes("jpeg")).decode('utf-8'))
    return base64_images

def extract_invoice_data(pdf_path: str) -> InvoiceData:
    base64_images = pdf_to_base64_images(pdf_path)
    content_array = [{"type": "text", "text": "Extract data from this multi-page invoice. If missing, leave null. Evaluate confidence ('High', 'Medium', 'Low'). Note page numbers."}]
    
    for img_base64 in base64_images:
        content_array.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}})

    return client.chat.completions.create(
        model=azure_deployment, 
        response_model=InvoiceData, 
        messages=[
            {"role": "system", "content": "You are an expert accountant. Extract data accurately and provide honest confidence scores."},
            {"role": "user", "content": content_array}
        ]
    )

# ==============================================================
# 4. Excel Setup Utility
# ==============================================================

def setup_excel_workbook():
    wb = openpyxl.Workbook()
    
    # --- Sheet 1: Details ---
    ws_details = wb.active
    ws_details.title = "Invoice Details"
    
    # Notice: Vendor Address and Date are back, and all Confidences are at the end
    details_headers = [
        "File Name", "Page #", "Vendor Name", "Vendor Address", "Bill To", "Ship To", "Remit To",
        "Origin", "Destination", "Invoice Number", "Date", "Currency", 
        "Material", "Description", "Quantity", "UOM", "Unit Price", "Line Total", "Invoice Total",
        "Inv# Conf", "Origin Conf", "Dest Conf", "UOM Conf", "Total Conf"
    ]
    ws_details.append(details_headers)
    
    # --- Sheet 2: QC Summary ---
    ws_qc = wb.create_sheet(title="QC Summary")
    qc_headers = [
        "File Name", "Invoice Number", "Status", "Extracted Total", 
        "Calculated Sum (Lines+Tax+Ship)", "Variance", "Currency"
    ]
    ws_qc.append(qc_headers)
    
    # Styling
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
# 5. Main Execution
# ==============================================================

if __name__ == "__main__":
    input_folder = r"C:\path\to\your\invoices"  # Folder containing all PDFs
    output_excel = os.path.join(input_folder, "Extracted_Invoices_Master.xlsx")
        
    wb, ws_details, ws_qc = setup_excel_workbook()
    red_font = Font(color="9C0006", bold=True)

    print(f"Scanning folder: {input_folder}")
    
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            print(f"\nProcessing: {filename}...")
            
            try:
                extracted_data = extract_invoice_data(pdf_path)
                
                # Math Validation
                calculated_line_sum = sum(item.line_total for item in extracted_data.line_items if item.line_total is not None)
                total_calculated = calculated_line_sum + extracted_data.tax_amount + extracted_data.shipping_handling
                variance = round(extracted_data.total_amount - total_calculated, 2)
                
                # Check Confidence
                confidence_fields = [
                    extracted_data.invoice_number_confidence, extracted_data.origin_confidence,
                    extracted_data.destination_confidence, extracted_data.total_amount_confidence
                ]
                for item in extracted_data.line_items:
                    if item.uom_confidence == "Low":
                        confidence_fields.append("Low")
                
                needs_review = (variance != 0.0) or ("Low" in confidence_fields)
                status = "FAIL - NEEDS REVIEW" if needs_review else "PASS"
                
                # -------------------------------------------------------------
                # Helper function to ensure Invoice Data repeats on EVERY row
                # -------------------------------------------------------------
                def create_row(page_num, material, desc, qty, uom, uom_conf, price, line_total):
                    return [
                        filename, page_num, extracted_data.vendor_name, extracted_data.vendor_address,
                        extracted_data.bill_to, extracted_data.ship_to, extracted_data.remit_to,
                        extracted_data.origin, extracted_data.destination, extracted_data.invoice_number,
                        extracted_data.date, extracted_data.currency, 
                        material, desc, qty, uom, price, line_total, extracted_data.total_amount,
                        extracted_data.invoice_number_confidence, extracted_data.origin_confidence, 
                        extracted_data.destination_confidence, uom_conf, extracted_data.total_amount_confidence
                    ]

                # 1. Write actual line items
                for item in extracted_data.line_items:
                    ws_details.append(create_row(
                        item.page_number, item.material, item.description, item.quantity, 
                        item.uom, item.uom_confidence, item.unit_price, item.line_total
                    ))
                    
                # 2. Write Shipping/Tax rows (repeating the invoice data)
                if extracted_data.shipping_handling > 0:
                    ws_details.append(create_row(None, None, "Shipping/Handling", None, None, None, None, extracted_data.shipping_handling))
                if extracted_data.tax_amount > 0:
                    ws_details.append(create_row(None, None, "Tax", None, None, None, None, extracted_data.tax_amount))

                # Write QC Summary
                ws_qc.append([
                    filename, extracted_data.invoice_number, status, 
                    extracted_data.total_amount, total_calculated, variance, extracted_data.currency
                ])
                
                if needs_review:
                    ws_qc.cell(row=ws_qc.max_row, column=3).font = red_font
                    print(f"⚠️ FLAG: Marked '{filename}' as NEEDS REVIEW in QC Summary.")
                else:
                    print(f"✅ PASS: {filename} processed perfectly.")
                    
            except Exception as e:
                print(f"❌ Error processing {filename}: {e}")
                
    # Auto-adjust column widths
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

    wb.save(output_excel)
    print(f"\n🎉 All finished! Data saved to: {output_excel}")
