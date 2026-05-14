import os
import re
import datetime
import pandas as pd
import instructor
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from dotenv import load_dotenv

load_dotenv()

# --- 1. SCHEMA ---
class LineItem(BaseModel):
    material: Optional[str] = Field(None, description="Material ID, Item Code, or SKU")
    description: Optional[str] = Field(None, description="Description of the item/service/fee")
    quantity: Optional[float] = Field(None)
    uom: Optional[str] = Field(None, description="Unit of Measure")
    unit_price: Optional[float] = Field(None)
    amount: Optional[float] = Field(None, description="Line total amount")

class UnifiedInvoice(BaseModel):
    supplier_name: Optional[str] = Field(None, description="Name of the issuing vendor/supplier")
    supplier_address: Optional[str] = Field(None, description="Full address of the supplier")
    invoice_number: Optional[str] = Field(None)
    invoice_date: Optional[str] = Field(None)
    remit_to: Optional[str] = Field(None, description="Full address where payment should be sent")
    shipper: Optional[str] = Field(None, description="Full address of the shipper")
    bill_to: Optional[str] = Field(None, description="Full address of the entity being billed")
    origin: Optional[str] = Field(None, description="Full starting/origin address")
    destination: Optional[str] = Field(None, description="Full destination/delivery address")
    subtotal: Optional[float] = Field(None)
    invoice_total: Optional[float] = Field(None, description="Grand total of the invoice")
    currency: Optional[str] = Field(None)
    line_items: List[LineItem]

# --- 2. CONFIDENCE RESOLVER ---
class ConfidenceMapper:
    def __init__(self, analyze_result):
        self.word_map = {}
        if analyze_result.pages:
            for page in analyze_result.pages:
                if page.words:
                    for word in page.words:
                        clean_text = word.content.strip().lower()
                        self.word_map[clean_text] = word.confidence

    def get_phrase_confidence(self, value):
        if value is None or str(value).strip() == "":
            return 0.0
        search_vals = str(value).strip().lower().split()
        confidences = [self.word_map.get(w, 0.5) for w in search_vals]
        return round(sum(confidences) / len(confidences), 2) if confidences else 0.5

# --- 3. EXTRACTION ENGINE ---
def process_invoice(file_path: str):
    file_name = os.path.basename(file_path)
    
    # Initialize Clients
    di_client = DocumentIntelligenceClient(
        endpoint=os.getenv("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT"),
        credential=AzureKeyCredential(os.getenv("AZURE_DOCUMENT_INTELLIGENCE_KEY"))
    )
    ai_client = instructor.from_openai(
        AzureOpenAI(
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            api_version="2024-02-15-preview"
        )
    )

    # Step 1: Azure Layout
    print(f"Status: Analyzing layout for {file_name}...")
    with open(file_path, "rb") as f:
        poller = di_client.begin_analyze_document(
            "prebuilt-layout",
            AnalyzeDocumentRequest(bytes_source=f.read()),
            output_content_format="markdown"
        )
    result = poller.result()
    conf_resolver = ConfidenceMapper(result)
    page_count = len(result.pages) if result.pages else 1

    # Step 2: Instructor Extraction with STRICT Negative Constraints
    print("Status: Extracting structured data via LLM...")
    system_prompt = (
        "You are a strict data extraction engine for a financial system. "
        "GUARDRAILS: This document may contain noise such as attached receipts. IGNORE all noise. "
        "RULE FOR CHARGES: Any additional fees (such as 'Tax', 'HST (On)', 'Freight') MUST be extracted "
        "as separate rows and appended to the `line_items` array (put fee name in 'description', value in 'amount'). "
        "CRITICAL RESTRICTION: DO NOT extract 'Subtotal', 'Total', 'Invoice Total', 'Amount Due', or 'Balance Due' "
        "as line item rows! These final summation amounts must strictly remain ONLY in the `subtotal` and `invoice_total` "
        "header fields. Do not add them to the `line_items` array. Do not infer or calculate missing fields."
    )

    extracted_data = ai_client.chat.completions.create(
        model="gpt-4o-mini",
        response_model=UnifiedInvoice,
        max_retries=3,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": result.content}
        ]
    )

    # Step 3: Archive the Markdown
    supplier_str = extracted_data.supplier_name or "Unknown_Supplier"
    safe_supplier = re.sub(r'[^\w\-_\. ]', '_', supplier_str).replace(' ', '_')
    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    
    archive_dir = "markdown_archive"
    os.makedirs(archive_dir, exist_ok=True)
    md_filename = os.path.join(archive_dir, f"{date_str}_{safe_supplier}.md")
    
    with open(md_filename, "w", encoding="utf-8") as md_file:
        md_file.write(result.content)

    # Step 4: Flatten the Data
    all_rows = []
    header_data = {"file_name": file_name, "pages": page_count}
    
    for field, value in extracted_data.model_dump(exclude={'line_items'}).items():
        header_data[field] = value
        header_data[f"{field}_conf"] = conf_resolver.get_phrase_confidence(value)

    if not extracted_data.line_items:
        all_rows.append(header_data)
    else:
        for item in extracted_data.line_items:
            row_data = header_data.copy()
            for field, value in item.model_dump().items():
                row_data[field] = value
                row_data[f"{field}_conf"] = conf_resolver.get_phrase_confidence(value)
            all_rows.append(row_data)

    df = pd.DataFrame(all_rows)

    # Step 5: COLUMN REORDERING AND RENAMING
    base_order = [
        "file_name", "supplier_name", "supplier_address", "invoice_number", 
        "invoice_date", "remit_to", "shipper", "bill_to", "origin", 
        "destination", "material", "description", "quantity", "uom", 
        "unit_price", "amount", "subtotal", "invoice_total", "currency", "pages"
    ]
    
    for col in base_order:
        if col not in df.columns:
            df[col] = None

    conf_cols = [col for col in df.columns if col.endswith('_conf')]
    df = df[base_order + conf_cols]
    
    rename_map = {
        "file_name": "File name", "supplier_name": "Supplier name", "supplier_address": "Supplier Address",
        "invoice_number": "Invoice number", "invoice_date": "Invoice date", "remit_to": "Remit To (full address)",
        "shipper": "Shipper(Full address)", "bill_to": "Bill to(full address)", "origin": "origin (full add)",
        "destination": "Destination,(full address)", "material": "Material", "description": "Description",
        "quantity": "Quantity", "uom": "UOM", "unit_price": "Unit Price", "amount": "Amount",
        "subtotal": "SubTotal", "invoice_total": "Invoice Total", "currency": "Currency", "pages": "Pages"
    }
    
    rename_conf_map = {f"{k}_conf": f"{v}_conf" for k, v in rename_map.items()}
    rename_map.update(rename_conf_map)
    df = df.rename(columns=rename_map)

    return df, extracted_data

# --- 6. EXECUTION ---
if __name__ == "__main__":
    FILE_PATH = "sample_invoice.pdf" 
    
    try:
        df, raw_obj = process_invoice(FILE_PATH)
        
        # Save to CSV
        safe_inv_num = raw_obj.invoice_number if raw_obj.invoice_number else "export"
        csv_filename = f"invoice_{safe_inv_num}.csv"
        df.to_csv(csv_filename, index=False)
        print(f"\n✅ Success! Data saved to: {csv_filename}")
        
    except Exception as e:
        print(f"Critical Error: {e}")
