import fitz  # PyMuPDF
import base64
import instructor
import csv
import os
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional

# ==============================================================
# 1. Define Data Schema (Pydantic)
# Instructor will force the LLM to output exactly this structure
# ==============================================================

class LineItem(BaseModel):
    description: str = Field(description="The name or description of the purchased item")
    quantity: Optional[float] = Field(None, description="The number of items purchased")
    unit_price: Optional[float] = Field(None, description="The price of a single unit")
    total: float = Field(description="The total cost for this line item")

class InvoiceData(BaseModel):
    vendor_name: str = Field(description="The name of the company issuing the invoice")
    invoice_number: Optional[str] = Field(None, description="The unique invoice ID or number")
    date: Optional[str] = Field(None, description="The date the invoice was issued")
    total_amount: float = Field(description="The final total amount charged on the invoice")
    line_items: List[LineItem] = Field(description="A list of all individual items purchased")

# ==============================================================
# 2. Setup Azure OpenAI Client with Instructor
# ==============================================================

# Ensure these environment variables are set in your system
azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
azure_api_key = os.getenv("AZURE_OPENAI_KEY")
azure_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT")
azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION")

# Patch the Azure client with Instructor
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
    """Converts a specific page of a PDF to a base64 encoded JPEG."""
    doc = fitz.open(pdf_path)
    page = doc[page_number]
    pix = page.get_pixmap(dpi=150)
    img_bytes = pix.tobytes("jpeg")
    return base64.b64encode(img_bytes).decode('utf-8')

# ==============================================================
# 4. LLM Data Extraction
# ==============================================================

def extract_invoice_data(pdf_path: str) -> InvoiceData:
    """Extracts structured data from an invoice PDF."""
    print(f"Processing image for {pdf_path}...")
    base64_image = pdf_page_to_base64(pdf_path, page_number=0)
    
    print("Calling Azure OpenAI API via Instructor...")
    
    # Instructor handles the function calling and validation automatically
    invoice_extraction = client.chat.completions.create(
        model=azure_deployment, 
        response_model=InvoiceData, # This is where the magic happens
        messages=[
            {
                "role": "system",
                "content": "You are an expert accountant. Extract the requested data from the provided invoice image. Strip away currency symbols and formatting."
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
        ],
        temperature=0.0 
    )
    return invoice_extraction

# ==============================================================
# 5. Save Data to CSV
# ==============================================================

def save_invoice_to_csv(invoice_data: InvoiceData, output_filename: str = "invoice_data.csv"):
    """Takes the structured InvoiceData object and flattens it into a CSV."""
    headers = [
        "Vendor Name", "Invoice Number", "Date", 
        "Item Description", "Quantity", "Unit Price", 
        "Item Total", "Invoice Total Amount"
    ]
    
    with open(output_filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        
        for item in invoice_data.line_items:
            row = [
                invoice_data.vendor_name, invoice_data.invoice_number,
                invoice_data.date, item.description, item.quantity,
                item.unit_price, item.total, invoice_data.total_amount
            ]
            writer.writerow(row)
            
    print(f"✅ Successfully saved structured data to {output_filename}")

# ==============================================================
# 6. Main Execution
# ==============================================================

if __name__ == "__main__":
    sample_pdf = "sample_invoice.pdf" 
    
    try:
        extracted_data = extract_invoice_data(sample_pdf)
        
        # Verify the raw JSON output
        print("\n--- Raw JSON Validated by Instructor ---")
        print(extracted_data.model_dump_json(indent=2))
        print("----------------------------------------\n")
        
        # Export to CSV
        save_invoice_to_csv(extracted_data, "extracted_invoice.csv")
        
    except FileNotFoundError:
        print(f"Error: Could not find '{sample_pdf}'. Please ensure the path is correct.")
    except Exception as e:
        print(f"An error occurred: {e}")
