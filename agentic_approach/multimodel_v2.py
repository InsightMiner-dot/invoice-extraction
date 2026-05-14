import os
import re
import datetime
import asyncio
import pandas as pd
import instructor
from openai import AsyncAzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence.aio import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from dotenv import load_dotenv

# Load credentials from .env file
load_dotenv()

# --- 1. SCHEMA (STRICTLY USER ALIASES ONLY) ---
class LineItem(BaseModel):
    material: Optional[str] = Field(None, description="Material ID, Item Code, or SKU")
    description: Optional[str] = Field(None, description="Description of the item, service, or fee")
    quantity: Optional[float] = Field(None, description="Quantity")
    uom: Optional[str] = Field(None, description="Unit of Measure")
    unit_price: Optional[float] = Field(None, description="Unit Price")
    amount: Optional[float] = Field(None, description="Line total amount")

class UnifiedInvoice(BaseModel):
    supplier_name: Optional[str] = Field(None, description="Name of the issuing vendor or supplier")
    supplier_address: Optional[str] = Field(None, description="Full complete supplier address")
    
    # EXACT ALIASES AS REQUESTED
    invoice_number: Optional[str] = Field(None, description="Unique Invoice number. Aliases: Inv no, Inv #")
    invoice_date: Optional[str] = Field(None, description="Invoice date")
    
    remit_to: Optional[str] = Field(None, description="Full complete 'Remit To' address")
    shipper: Optional[str] = Field(None, description="Full complete 'Shipper' address")
    bill_to: Optional[str] = Field(None, description="Full complete 'Bill To' address")
    
    # EXACT ALIASES AS REQUESTED
    origin: Optional[str] = Field(None, description="Full complete 'Origin' address. Aliases: Ship From, Pickup, Generator")
    destination: Optional[str] = Field(None, description="Full complete 'Destination' address. Aliases: Consignee, Deliver To, Designated")
    
    subtotal: Optional[float] = Field(None, description="Subtotal before tax")
    invoice_total: Optional[float] = Field(None, description="Grand total of the invoice")
    currency: Optional[str] = Field(None, description="Currency code (e.g., USD, CAD)")
    
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

# --- 3. ASYNC EXTRACTION ENGINE ---
async def process_single_invoice_async(file_path: str, di_client: DocumentIntelligenceClient, ai_client, semaphore: asyncio.Semaphore):
    """Processes a single invoice asynchronously to enable fast batch processing."""
    
    async with semaphore:
        file_name = os.path.basename(file_path)
        print(f"[{file_name}] Started processing...")
        
        try:
            # Step 1: Azure Layout (Async)
            with open(file_path, "rb") as f:
                poller = await di_client.begin_analyze_document(
                    "prebuilt-layout",
                    AnalyzeDocumentRequest(bytes_source=f.read()),
                    output_content_format="markdown"
                )
            result = await poller.result()
            conf_resolver = ConfidenceMapper(result)
            page_count = len(result.pages) if result.pages else 1

            # Step 2: Instructor Extraction (Async)
            system_prompt = (
                "You are a strict data extraction engine for a financial system. "
                "GUARDRAILS: This document may contain noise such as attached receipts. IGNORE all noise. "
                "MULTI-PAGE SCANNING: You must thoroughly scan ALL pages of the provided document text to find "
                "the required fields, especially Origin, Destination, and other addresses. They may be located "
                "at the very end of a multi-page document. "
                "RULE FOR CHARGES: Any additional fees (such as 'Tax', 'HST (On)', 'Freight') MUST be extracted "
                "as separate rows and appended to the `line_items` array (put fee name in 'description', value in 'amount'). "
                "CRITICAL RESTRICTION: DO NOT extract 'Subtotal', 'Total', 'Invoice Total', 'Amount Due', or 'Balance Due' "
                "as line item rows! These final summation amounts must strictly remain ONLY in the `subtotal` and `invoice_total` "
                "header fields. Do not add them to the `line_items` array. Do not infer or calculate missing fields."
            )

            extracted_data = await ai_client.chat.completions.create(
                model="gpt-4o-mini", # Note: For highly complex multi-page files, gpt-4o performs cross-page scanning better than mini
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
            md_filename = os.path.join(archive_dir, f"{date_str}_{safe_supplier}_{file_name}.md")
            
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
                "invoice_number": "Invoice number", "invoice_date": "Invoice date", 
                "remit_to": "Remit To", "shipper": "Shipper", "bill_to": "Bill To", 
                "origin": "Origin", "destination": "Destination", 
                "material": "Material", "description": "Description",
                "quantity": "Quantity", "uom": "UOM", "unit_price": "Unit Price", "amount": "Amount",
                "subtotal": "SubTotal", "invoice_total": "Invoice Total", "currency": "Currency", "pages": "Pages"
            }
            
            rename_conf_map = {f"{k}_conf": f"{v}_conf" for k, v in rename_map.items()}
            rename_map.update(rename_conf_map)
            df = df.rename(columns=rename_map)

            print(f"[{file_name}] ✅ Successfully extracted {len(all_rows)} rows.")
            return df
            
        except Exception as e:
            print(f"[{file_name}] ❌ Failed: {str(e)}")
            return None


# --- 4. BATCH PROCESSOR ORCHESTRATION ---
async def process_folder_batch(input_folder: str, output_csv: str):
    print(f"\n--- Starting Batch Processing for folder: {input_folder} ---")
    
    # 1. Gather Supported Files
    supported_extensions = ('.pdf', '.png', '.jpg', '.jpeg', '.tiff')
    files_to_process = [
        os.path.join(input_folder, f) for f in os.listdir(input_folder) 
        if f.lower().endswith(supported_extensions)
    ]
    
    if not files_to_process:
        print("No valid files found in the directory.")
        return

    print(f"Found {len(files_to_process)} documents. Initializing async clients...")

    # 2. Initialize Async Clients
    di_client = DocumentIntelligenceClient(
        endpoint=os.getenv("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT"),
        credential=AzureKeyCredential(os.getenv("AZURE_DOCUMENT_INTELLIGENCE_KEY"))
    )
    
    ai_client = instructor.from_openai(
        AsyncAzureOpenAI(
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            api_version="2024-02-15-preview"
        )
    )

    # 3. Setup Concurrency Limit
    semaphore = asyncio.Semaphore(5) 
    
    # 4. Fire Async Tasks
    tasks = [
        process_single_invoice_async(file_path, di_client, ai_client, semaphore)
        for file_path in files_to_process
    ]
    
    # 5. Gather Results
    results = await asyncio.gather(*tasks)
    
    # 6. Merge & Export
    valid_dfs = [df for df in results if df is not None]
    
    if valid_dfs:
        master_df = pd.concat(valid_dfs, ignore_index=True)
        master_df.to_csv(output_csv, index=False)
        print(f"\n✅ Batch Complete! Extracted a total of {len(master_df)} rows.")
        print(f"✅ Master CSV saved to: {output_csv}")
    else:
        print("\n❌ Batch failed: No data could be extracted from any documents.")
        
    await di_client.close()

# --- 5. EXECUTION ENTRY POINT ---
if __name__ == "__main__":
    # Point this to your folder full of invoices
    INPUT_DIRECTORY = "./invoice_batch"  
    MASTER_OUTPUT_CSV = "master_extracted_invoices.csv"
    
    os.makedirs(INPUT_DIRECTORY, exist_ok=True)
    
    # Execute the async event loop
    asyncio.run(process_folder_batch(INPUT_DIRECTORY, MASTER_OUTPUT_CSV))
