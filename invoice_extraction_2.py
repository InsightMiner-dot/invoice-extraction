import fitz  # PyMuPDF
import base64
import instructor
import csv
import os
import shutil
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional

# ==============================================================
# 1. Define Data Schema (QC, CONFIDENCE, & PAGE #)
# ==============================================================

class LineItem(BaseModel):
    page_number: Optional[int] = Field(None, description="The page number (starting at 1) where this item was found")
    material: Optional[str] = Field(None, description="The material code, part number, or SKU")
    description: str = Field(description="The name or description of the item")
    quantity: Optional[float] = Field(None, description="The number of items purchased")
    
    uom: Optional[str] = Field(None, description="Unit of Measure (e.g., EA, LBS, KG)")
    uom_confidence: Optional[str] = Field(None, description="Rate confidence in UOM extraction: 'High', 'Medium', or 'Low'")
    
    unit_price: Optional[float] = Field(None, description="The price of a single unit")
    line_total: float = Field(description="The total cost for this specific line item. Look for 'Amount', 'Total', or 'Extended Price'.")

class InvoiceData(BaseModel):
    vendor_name: str = Field(description="The name of the company issuing the invoice")
    
    invoice_number: Optional[str] = Field(None, description="The unique invoice ID or number")
    invoice_number_confidence: Optional[str] = Field(None, description="Rate confidence in Invoice Number extraction: 'High', 'Medium', or 'Low'")
    
    date: Optional[str] = Field(None, description="The date the invoice was issued")
    vendor_address: Optional[str] = Field(None, description="The full address of the vendor")
    bill_to: Optional[str] = Field(None, description="The full 'Bill To' or 'Sold To' address")
    ship_to: Optional[str] = Field(None, description="The 'Ship To' address.")
    remit_to: Optional[str] = Field(None, description="The 'Remit To' address")
    
    origin: Optional[str] = Field(None, description="The origin address ('From', 'Ship From', 'Pickup', 'Generator').")
    origin_confidence: Optional[str] = Field(None, description="Rate confidence in Origin extraction: 'High', 'Medium', or 'Low'")
    
    destination: Optional[str] = Field(None, description="The destination address ('To', 'Deliver To', 'Consignee').")
    destination_confidence: Optional[str] = Field(None, description="Rate confidence in Destination extraction: 'High', 'Medium', or 'Low'")
    
    currency: Optional[str] = Field(None, description="The 3-letter currency code (e.g., USD, CAD, EUR)")
    
    tax_amount: Optional[float] = Field(0.0, description="The total tax amount. Default to 0.0 if not explicitly found.")
    shipping_handling: Optional[float] = Field(0.0, description="Shipping/Freight charges. Default to 0.0 if not explicitly found.")
    
    total_amount: float = Field(description="The final total amount charged on the invoice")
    total_amount_confidence: Optional[str] = Field(None, description="Rate confidence in Total Amount extraction: 'High', 'Medium', or 'Low'")
    
    line_items: List[LineItem] = Field(description="A list of all individual items purchased across all pages")

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
# 3. PDF to Image Conversion 
# ==============================================================

def pdf_to_base64_images(pdf_path: str) -> List[str]:
    doc = fitz.open(pdf_path)
    base64_images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        base64_images.append(base64.b64encode(pix.tobytes("jpeg")).decode('utf-8'))
    return base64_images

# ==============================================================
# 4. LLM Data Extraction 
# ==============================================================

def extract_invoice_data(pdf_path: str) -> InvoiceData:
    print(f"Converting pages of {os.path.basename(pdf_path)} to images...")
    base64_images = pdf_to_base64_images(pdf_path)
    
    content_array = [{"type": "text", "text": "Extract data from this multi-page invoice. If a field is missing, leave it null. Evaluate your confidence ('High', 'Medium', 'Low') for the requested fields based on image clarity. Note the page number (1, 2, etc.) for each line item."}]
    
    for img_base64 in base64_images:
        content_array.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}})

    print("Extracting data via Azure OpenAI...")
    return client.chat.completions.create(
        model=azure_deployment, 
        response_model=InvoiceData, 
        messages=[
            {"role": "system", "content": "You are an expert accountant. Extract data accurately, handle varied address aliases, and provide honest confidence scores."},
            {"role": "user", "content": content_array}
        ]
    )

# ==============================================================
# 5. Save Detail Data to CSV
# ==============================================================

def save_invoice_to_csv(invoice_data: InvoiceData, pdf_path: str, output_filename: str = "invoice_details.csv"):
    file_name = os.path.basename(pdf_path)
    
    headers = [
        "File Name", "Page #", "Vendor Name", "Invoice Number", "Inv# Confidence",
        "Origin", "Origin Conf", "Destination", "Dest Conf", "Bill To", "Ship To", "Remit To",
        "Currency", "Material", "Description", "Quantity", "UOM", "UOM Conf", 
        "Unit Price", "Line Total", "Invoice Total", "Total Conf"
    ]
    
    file_exists = os.path.isfile(output_filename)
    with open(output_filename, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(headers)
            
        for item in invoice_data.line_items:
            writer.writerow([
                file_name, item.page_number, invoice_data.vendor_name, 
                invoice_data.invoice_number, invoice_data.invoice_number_confidence,
                invoice_data.origin, invoice_data.origin_confidence,
                invoice_data.destination, invoice_data.destination_confidence,
                invoice_data.bill_to, invoice_data.ship_to, invoice_data.remit_to,
                invoice_data.currency, item.material, item.description, 
                item.quantity, item.uom, item.uom_confidence, item.unit_price, 
                item.line_total, invoice_data.total_amount, invoice_data.total_amount_confidence
            ])
            
        # Append Shipping if exists
        if invoice_data.shipping_handling > 0:
            writer.writerow([file_name, "", invoice_data.vendor_name, invoice_data.invoice_number, "", "", "", "", "", "", "", "", invoice_data.currency, "", "Shipping/Handling", "", "", "", "", invoice_data.shipping_handling, invoice_data.total_amount, ""])

        # Append Tax if exists
        if invoice_data.tax_amount > 0:
            writer.writerow([file_name, "", invoice_data.vendor_name, invoice_data.invoice_number, "", "", "", "", "", "", "", "", invoice_data.currency, "", "Tax", "", "", "", "", invoice_data.tax_amount, invoice_data.total_amount, ""])

# ==============================================================
# 6. Save QC Math Verification to CSV
# ==============================================================

def save_qc_summary(invoice_data: InvoiceData, pdf_path: str, output_filename: str = "qc_summary.csv"):
    file_name = os.path.basename(pdf_path)
    
    calculated_line_sum = sum(item.line_total for item in invoice_data.line_items if item.line_total is not None)
    total_calculated = calculated_line_sum + invoice_data.tax_amount + invoice_data.shipping_handling
    variance = round(invoice_data.total_amount - total_calculated, 2)
    
    status = "PASS" if variance == 0.0 else "FAIL - NEEDS REVIEW"
    
    headers = [
        "File Name", "Invoice Number", "Status", "Extracted Total", 
        "Calculated Sum (Lines + Tax + Ship)", "Variance", "Currency"
    ]
    
    file_exists = os.path.isfile(output_filename)
    with open(output_filename, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(headers)
            
        writer.writerow([
            file_name, invoice_data.invoice_number, status, 
            invoice_data.total_amount, total_calculated, variance, invoice_data.currency
        ])

# ==============================================================
# 7. Main Execution & Auto-Routing
# ==============================================================

if __name__ == "__main__":
    sample_pdf = r"sample_invoice.pdf" 
    review_folder = "Needs_Manual_Review"
    
    if not os.path.exists(review_folder):
        os.makedirs(review_folder)
    
    try:
        # 1. Extract Data
        extracted_data = extract_invoice_data(sample_pdf)
        
        # 2. Save Outputs
        save_invoice_to_csv(extracted_data, sample_pdf, "invoice_details.csv")
        save_qc_summary(extracted_data, sample_pdf, "qc_summary.csv")
        
        # 3. Validation & Routing
        calculated_line_sum = sum(item.line_total for item in extracted_data.line_items if item.line_total is not None)
        total_calculated = calculated_line_sum + extracted_data.tax_amount + extracted_data.shipping_handling
        variance = round(extracted_data.total_amount - total_calculated, 2)
        
        confidence_fields = [
            extracted_data.invoice_number_confidence,
            extracted_data.origin_confidence,
            extracted_data.destination_confidence,
            extracted_data.total_amount_confidence
        ]
        
        # Check line item UOM confidence as well
        for item in extracted_data.line_items:
            if item.uom_confidence == "Low":
                confidence_fields.append("Low")
        
        if variance != 0.0 or "Low" in confidence_fields:
            print(f"⚠️ FLAG TRIGGERED: Routing {os.path.basename(sample_pdf)} to '{review_folder}'")
            shutil.copy(sample_pdf, os.path.join(review_folder, os.path.basename(sample_pdf)))
        else:
            print(f"✅ Auto-Processing Complete: {os.path.basename(sample_pdf)} looks perfect.")

    except FileNotFoundError:
        print(f"Error: Could not find '{sample_pdf}'. Please verify the path.")
    except Exception as e:
        print(f"An error occurred: {e}")
