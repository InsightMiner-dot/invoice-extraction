import fitz  # PyMuPDF
import base64
import instructor
import csv
import os
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional

# ==============================================================
# 1. Define Data Schema (UPDATED WITH ADDRESSES & TAX)
# ==============================================================

class LineItem(BaseModel):
    description: str = Field(description="The name or description of the purchased item")
    quantity: Optional[float] = Field(None, description="The number of items purchased")
    unit_price: Optional[float] = Field(None, description="The price of a single unit")
    total: float = Field(description="The total cost for this line item")

class InvoiceData(BaseModel):
    # Standard Fields
    vendor_name: str = Field(description="The name of the company issuing the invoice")
    invoice_number: Optional[str] = Field(None, description="The unique invoice ID or number")
    date: Optional[str] = Field(None, description="The date the invoice was issued")
    
    # New Address Fields
    vendor_address: Optional[str] = Field(None, description="The full address of the vendor issuing the invoice")
    bill_to: Optional[str] = Field(None, description="The full 'Bill To' address")
    ship_to: Optional[str] = Field(None, description="The full 'Ship To' address")
    origin: Optional[str] = Field(None, description="The origin address where goods/services shipped from")
    destination: Optional[str] = Field(None, description="The destination address where goods/services are going")
    
    # Totals and Taxes
    tax_amount: Optional[float] = Field(None, description="The total tax amount charged on the invoice. Leave null if no tax is listed.")
    total_amount: float = Field(description="The final total amount charged on the invoice")
    
    # Items
    line_items: List[LineItem] = Field(description="A list of all individual items purchased")

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
# 3. PDF to Image Conversion
# ==============================================================

def pdf_page_to_base64(pdf_path: str, page_number: int = 0) -> str:
    doc = fitz.open(pdf_path)
    page = doc[page_number]
    pix = page.get_pixmap(dpi=150)
    img_bytes = pix.tobytes("jpeg")
    return base64.b64encode(img_bytes).decode('utf-8')

# ==============================================================
# 4. LLM Data Extraction
# ==============================================================

def extract_invoice_data(pdf_path: str) -> InvoiceData:
    print(f"Processing image for {pdf_path}...")
    base64_image = pdf_page_to_base64(pdf_path, page_number=0)
    
    print("Calling Azure OpenAI API via Instructor...")
    
    invoice_extraction = client.chat.completions.create(
        model=azure_deployment, 
        response_model=InvoiceData, 
        messages=[
            {
                "role": "system",
                "content": "You are an expert accountant and logistics coordinator. Extract the requested data from the provided invoice image. Strip away currency symbols."
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text", 
                        "text": "Please extract the data from this invoice. If a field is missing, leave it null."
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}"
                        }
                    }
                ]
            }
        ]
        # Temperature omitted for reasoning models
    )
    return invoice_extraction

# ==============================================================
# 5. Save Data to CSV (UPDATED TO HANDLE ADDRESSES & TAX ROW)
# ==============================================================

def save_invoice_to_csv(invoice_data: InvoiceData, pdf_path: str, output_filename: str = "invoice_data.csv"):
    file_name = os.path.basename(pdf_path)
    
    # We added all the new address columns here
    headers = [
        "File Name", "Vendor Name", "Vendor Address", "Bill To", "Ship To", 
        "Origin", "Destination", "Invoice Number", "Date", 
        "Item Description", "Quantity", "Unit Price", 
        "Item Total", "Invoice Total Amount"
    ]
    
    with open(output_filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        
        # A quick helper function so we don't have to rewrite the address variables over and over
        def create_row(description, quantity, unit_price, item_total):
            return [
                file_name, 
                invoice_data.vendor_name, 
                invoice_data.vendor_address,
                invoice_data.bill_to,
                invoice_data.ship_to,
                invoice_data.origin,
                invoice_data.destination,
                invoice_data.invoice_number,
                invoice_data.date, 
                description, 
                quantity, 
                unit_price, 
                item_total, 
                invoice_data.total_amount
            ]
        
        # 1. Write the standard line items
        for item in invoice_data.line_items:
            writer.writerow(create_row(item.description, item.quantity, item.unit_price, item.total))
            
        # 2. Write the Tax as its own line item at the end (if it exists)
        if invoice_data.tax_amount is not None and invoice_data.tax_amount > 0:
            writer.writerow(create_row("Tax", None, None, invoice_data.tax_amount))
            
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
