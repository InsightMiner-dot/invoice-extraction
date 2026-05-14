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

# --- OVERRIDE ENV VARIABLES ---
load_dotenv(override=True)

# --- 1. SCHEMA (STRICTLY USER ALIASES + ROW INDEPENDENCE) ---
class LineItem(BaseModel):
    material: Optional[str] = Field(None, description="Material ID, Item Code, or SKU")
    description: Optional[str] = Field(None, description="Description of the item, service, or fee")
    quantity: Optional[float] = Field(None, description="Quantity")
    uom: Optional[str] = Field(None, description="Unit of Measure. Do NOT carry over from previous rows. Return NULL if missing on this specific line.")
    unit_price: Optional[float] = Field(None, description="Unit Price")
    amount: Optional[float] = Field(None, description="Line total amount")

class UnifiedInvoice(BaseModel):
    supplier_name: Optional[str] = Field(None, description="Name of the issuing vendor or supplier")
    supplier_address: Optional[str] = Field(None, description="Full complete supplier address")
    
    invoice_number: Optional[str] = Field(None, description="Unique Invoice number. Aliases: Inv no, Inv #")
    invoice_date: Optional[str] = Field(None, description="Invoice date")
    
    remit_to: Optional[str] = Field(None, description="Full complete 'Remit To' address")
    shipper: Optional[str] = Field(None, description="Full complete 'Shipper' address")
    bill_to: Optional[str] = Field(None, description="Full complete 'Bill To' address")
    
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

# --- 3. QUALITY CONTROL ENGINE ---
def run_qc_checks(invoice_obj: UnifiedInvoice, conf_mapper: ConfidenceMapper) -> dict:
    reasons = []
    status = "PASS"

    # A. Critical Field Check
    if not invoice_obj.invoice_number: reasons.append("Missing Invoice Number")
    if not invoice_obj.origin: reasons.append("Missing Origin Address")
    if not invoice_obj.destination: reasons.append("Missing Destination Address")

    # B. AI Confidence Check (Threshold 85%)
    threshold = 0.85
    if invoice_obj.origin:
        conf = conf_mapper.get_phrase_confidence(invoice_obj.origin)
        if conf < threshold: reasons.append(f"Low Origin Confidence ({conf})")
            
    if invoice_obj.destination:
        conf = conf_mapper.get_phrase_confidence(invoice_obj.destination)
        if conf < threshold: reasons.append(f"Low Destination Confidence ({conf})")

    # C. Math Reconciliation Check
    if invoice_obj.invoice_total and invoice_obj.line_items:
        calc_total = sum([item.amount for item in invoice_obj.line_items if item.amount])
        # Margin of error of 0.05 for rounding differences
        if abs(calc_total - invoice_obj.invoice_total) > 0.05:
            reasons.append(f"Math Error: Lines sum to {calc_total}, Total is {invoice_obj.invoice_total}")
    elif not invoice_obj.invoice_total:
        reasons.append("Missing Grand Total for Math Check")

    if reasons:
        status = "REVIEW"

    return {
        "qc_status": status,
        "qc_reasons": " | ".join(reasons)
    }

# --- 4. ASYNC EXTRACTION ENGINE ---
async def process_single_invoice_async(file_path: str, di_client: DocumentIntelligenceClient, ai_client, semaphore: asyncio.Semaphore):
    
    async with semaphore:
        file_name = os.path.basename(file_path)
        print(f"[{file_name}] Started processing...")
        
        try:
            # 1. Azure Layout
            with open(file_path, "rb") as f:
                poller = await di_client.begin_analyze_document(
                    "prebuilt-layout",
                    AnalyzeDocumentRequest(bytes_source=f.read()),
                    output_content_format="markdown"
                )
            result = await poller.result()
            conf_resolver = ConfidenceMapper(result)
            page_count = len(result.pages) if result.pages else 1

            # 2. Instructor LLM
            system_prompt = (
                "You are a strict data extraction engine for a financial system. "
                "GUARDRAILS: This document may contain noise such as attached receipts. IGNORE all noise. "
                "MULTI-PAGE SCANNING: You must thoroughly scan ALL pages to find required fields, especially "
                "Origin, Destination, and other addresses. They may be located at the end of the document. "
                "ROW INDEPENDENCE: Treat every line item independently. DO NOT carry over, copy, or inherit values "
                "(especially UOM or Quantities) from previous rows. If not explicitly on that line, return NULL. "
                "RULE FOR CHARGES: Additional fees ('Tax', 'Freight') MUST be extracted as separate rows. "
                "CRITICAL RESTRICTION: DO NOT extract 'Subtotal', 'Total', 'Invoice Total' as line item rows! "
                "These summation amounts must strictly remain ONLY in the header fields."
            )

            extracted_data = await ai_client.chat.completions.create(
                model="gpt-4o-mini", 
                response_model=UnifiedInvoice,
                max_retries=3,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": result.content}
                ]
            )
            
            # 3. Run Quality Control
            qc_results = run_qc_checks(extracted_data, conf_resolver)
            
            # Create a summary row for this specific invoice
            summary_row = {
                "File Name": file_name,
                "Invoice Number": extracted_data.invoice_number,
                "Supplier Name": extracted_data.supplier_name,
                "QC Status": qc_results["qc_status"],
                "QC Reasons": qc_results["qc_reasons"]
            }

            # 4. Archive the Markdown
            supplier_str = extracted_data.supplier_name or "Unknown_Supplier"
            safe_supplier = re.sub(r'[^\w\-_\. ]', '_', supplier_str).replace(' ', '_')
            date_str = datetime.datetime.now().strftime("%Y-%m-%d")
            
            archive_dir = "markdown_archive"
            os.makedirs(archive_dir, exist_ok=True)
            md_filename = os.path.join(archive_dir, f"{date_str}_{safe_supplier}_{file_name}.md")
            
            with open(md_filename, "w", encoding="utf-8") as md_file:
                md_file.write(result.content)

            # 5. Flatten the Details Data
            all_rows = []
            header_data = {"file_name": file_name, "pages": page_count}
            
            for field, value in extracted_data.model_dump(exclude={'line_items'}).items():
                header_data[field] = value
                header_data[f"{field}_conf"] = conf_resolver.get_phrase_confidence(value)

            # Append QC data so it's also available on the detailed rows if needed
            header_data["qc_status"] = qc_results["qc_status"]
            
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

            # 6. Column Formatting
            base_order = [
                "file_name", "qc_status", "supplier_name", "supplier_address", "invoice_number", 
                "invoice_date", "remit_to", "shipper", "bill_to", "origin", 
                "destination", "material", "description", "quantity", "uom", 
                "unit_price", "amount", "subtotal", "invoice_total", "currency", "pages"
            ]
            for col in base_order:
                if col not in df.columns: df[col] = None

            conf_cols = [col for col in df.columns if col.endswith('_conf')]
            df = df[base_order + conf_cols]
            
            rename_map = {
                "file_name": "File name", "qc_status": "QC Status", "supplier_name": "Supplier name", 
                "supplier_address": "Supplier Address", "invoice_number": "Invoice number", 
                "invoice_date": "Invoice date", "remit_to": "Remit To", "shipper": "Shipper", 
                "bill_to": "Bill To", "origin": "Origin", "destination": "Destination", 
                "material": "Material", "description": "Description", "quantity": "Quantity", 
                "uom": "UOM", "unit_price": "Unit Price", "amount": "Amount",
                "subtotal": "SubTotal", "invoice_total": "Invoice Total", "currency": "Currency", "pages": "Pages"
            }
            
            rename_conf_map = {f"{k}_conf": f"{v}_conf" for k, v in rename_map.items()}
            rename_map.update(rename_conf_map)
            df = df.rename(columns=rename_map)

            print(f"[{file_name}] ✅ Finished. Status: {qc_results['qc_status']}")
            
            # RETURN BOTH THE DATAFRAME AND THE SUMMARY ROW
            return df, summary_row
            
        except Exception as e:
            print(f"[{file_name}] ❌ Failed: {str(e)}")
            return None


# --- 5. BATCH PROCESSOR ORCHESTRATION ---
async def process_folder_batch(input_folder: str, output_excel: str):
    print(f"\n--- Starting Batch Processing for folder: {input_folder} ---")
    
    supported_extensions = ('.pdf', '.png', '.jpg', '.jpeg', '.tiff')
    files_to_process = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.lower().endswith(supported_extensions)]
    
    if not files_to_process:
        print("No valid files found in the directory.")
        return

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

    semaphore = asyncio.Semaphore(5) 
    
    tasks = [process_single_invoice_async(f, di_client, ai_client, semaphore) for f in files_to_process]
    results = await asyncio.gather(*tasks)
    
    # Separate the Details DataFrames and the Summary Rows
    valid_dfs = [res[0] for res in results if res is not None]
    valid_summaries = [res[1] for res in results if res is not None]
    
    if valid_dfs:
        master_df = pd.concat(valid_dfs, ignore_index=True)
        summary_df = pd.DataFrame(valid_summaries)
        
        # --- EXCEL MULTI-SHEET EXPORT ---
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # Save the Summary to the first sheet
            summary_df.to_excel(writer, sheet_name='QC_Summary', index=False)
            # Save the Details to the second sheet
            master_df.to_excel(writer, sheet_name='Line_Items_Details', index=False)
            
        print(f"\n✅ Batch Complete! Extracted {len(master_df)} line items across {len(summary_df)} invoices.")
        print(f"✅ Master Multi-Sheet Excel saved to: {output_excel}")
        
        # Print a quick terminal summary of the QC
        pass_count = len(summary_df[summary_df['QC Status'] == 'PASS'])
        review_count = len(summary_df[summary_df['QC Status'] == 'REVIEW'])
        print(f"📊 QC Report: {pass_count} Passed STP | {review_count} Flagged for Human Review")
    else:
        print("\n❌ Batch failed: No data could be extracted.")
        
    await di_client.close()

# --- 6. EXECUTION ENTRY POINT ---
if __name__ == "__main__":
    INPUT_DIRECTORY = "./invoice_batch"  
    # Notice we changed the output to .xlsx
    MASTER_OUTPUT_EXCEL = "master_extracted_invoices.xlsx" 
    
    os.makedirs(INPUT_DIRECTORY, exist_ok=True)
    asyncio.run(process_folder_batch(INPUT_DIRECTORY, MASTER_OUTPUT_EXCEL))
