import fitz  # PyMuPDF
import base64
import instructor
import csv
import os
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional

# ==============================================================
# 1. Define Data Schema 
# ==============================================================

class LineItem(BaseModel):
    material: Optional[str] = Field(None, description="The material code, part number, SKU, or material type if specified")
    description: str = Field(description="The name or description of the purchased item")
    quantity: Optional[float] = Field(None, description="The number of items purchased")
    unit_price: Optional[float] = Field(None, description="The price of a single unit")
    line_total: float = Field(description="The total cost for this specific line item. Look for columns labeled 'Line Total', 'Amount', 'Total', or 'Extended Price'.")

class InvoiceData(BaseModel):
    vendor_name: str = Field(description="The name of the company issuing the invoice")
    invoice_number: Optional[str] = Field(None, description="The unique invoice ID or number")
    date: Optional[str] = Field(None, description="The date the invoice was issued")
    
    vendor_address: Optional[str] = Field(None, description="The full address of the vendor issuing the invoice")
    bill_to: Optional[str] = Field(None, description="The full 'Bill To' address")
    ship_to: Optional[str] = Field(None, description="The full 'Ship To' address")
    remit_to: Optional[str] = Field(None, description="The 'Remit To' address where payment should actually be sent")
    origin: Optional[str] = Field(None, description="The origin address where goods/services shipped from")
    destination: Optional[str] = Field(None, description="The destination address where goods/services are going")
    
    currency: Optional[str] = Field(None, description="The 3-letter currency code (e.g., USD, CAD, EUR) used on the invoice")
    
    tax_amount: Optional[float] = Field(None, description="The total tax amount charged on the invoice. Leave null if no tax is listed.")
    total_amount: float = Field(description="The final total amount charged on the invoice")
    
    line_items: List[LineItem] = Field(description="A list of all individual items purchased across all pages")

# ==============================================================
# 2. Setup Azure OpenAI Client with Instructor
# ==============================================================

azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
azure_api_key = os.getenv("AZURE_OPENAI_KEY")
azure_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT")
azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION")

client = instructor.from_openai(
    AzureOpenAI(
        azure_endpoint=azure_endpoint,
        api_key=azure_api_key,
        api_version=azure_api_version,
    )
)

# ==============================================================
# 3. PDF to Image Conversion (UPDATED FOR MULTIPLE PAGES)
# ==============================================================

def pdf_to_base64_images(pdf_path: str) -> List[str]:
    """Converts ALL pages of a PDF into a list of base64 encoded JPEGs."""
    doc = fitz.open(pdf_path)
    base64_images = []
    
    # Loop through every page in the document
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("jpeg")
        base64_images.append(base64.b64encode(img_bytes).decode('utf-8'))
        
    return base64_images

# ==============================================================
# 4. LLM Data Extraction (UPDATED FOR MULTIPLE PAGES)
# ==============================================================

def extract_invoice_data(pdf_path: str) -> InvoiceData:
    print(f"Converting all pages of {pdf_path} to images...")
    base64_images = pdf_to_base64_images(pdf_path)
    print(f"Found {len(base64_images)} page(s).")
    
    # Build the dynamic content array starting with our text prompt
    content_array = [
        {
            "type": "text", 
            "text": "Please extract the data from this multi-page invoice. If a field is missing, leave it null. Ensure you capture line items from all pages."
        }
    ]
    
    # Dynamically attach every page of the PDF to the prompt
    for img_base64 in base64_images:
        content_array.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/jpeg;base64,{img_base64}"
            }
        })

    print("Calling Azure OpenAI API via Instructor...")
    invoice_extraction = client.chat.completions.create(
        model=azure_deployment, 
        response_model=InvoiceData, 
        messages=[
            {
                "role": "system",
                "content": "You are an expert accountant and logistics coordinator. Extract the requested data from the provided invoice images. Strip away currency symbols from the number amounts, but note the actual currency in the currency field."
            },
            {
                "role": "user",
                "content": content_array # Pass the fully built array of images here
            }
        ]
    )
    return invoice_extraction

# ==============================================================
# 5. Save Data to CSV
# ==============================================================

def save_invoice_to_csv(invoice_data: InvoiceData, pdf_path: str, output_filename: str = "invoice_data.csv"):
    file_name = os.path.basename(pdf_path)
    
    headers = [
        "File Name", "Vendor Name", "Vendor Address", "Bill To", "Ship To", "Remit To",
        "Origin", "Destination", "Invoice Number", "Date", "Currency",
        "Material", "Item Description", "Quantity", "Unit Price", 
        "Line Total", "Invoice Total Amount"
    ]
    
    with open(output_filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        
        def create_row(material, description, quantity, unit_price, line_total):
            return [
                file_name, invoice_data.vendor_name, invoice_data.vendor_address,
                invoice_data.bill_to, invoice_data.ship_to, invoice_data.remit_to, 
                invoice_data.origin, invoice_data.destination, invoice_data.invoice_number,
                invoice_data.date, invoice_data.currency, material, 
                description, quantity, unit_price, line_total, invoice_data.total_amount
            ]
        
        for item in invoice_data.line_items:
            writer.writerow(create_row(item.material, item.description, item.quantity, item.unit_price, item.line_total))
            
        if invoice_data.tax_amount is not None and invoice_data.tax_amount > 0:
            writer.writerow(create_row(None, "Tax", None, None, invoice_data.tax_amount))
            
    print(f"✅ Successfully saved structured data to {output_filename}")

# ==============================================================
# 6. Main Execution
# ==============================================================

if __name__ == "__main__":
    sample_pdf = "sample_invoice.pdf" 
    
    try:
        extracted_data = extract_invoice_data(sample_pdf)
        
        print("\n--- Raw JSON Validated by Instructor ---")
        print(extracted_data.model_dump_json(indent=2))
        print("----------------------------------------\n")
        
        save_invoice_to_csv(extracted_data, sample_pdf, "extracted_invoice.csv")
        
    except FileNotFoundError:
        print(f"Error: Could not find '{sample_pdf}'. Please ensure the path is correct.")
    except Exception as e:
        print(f"An error occurred: {e}")
