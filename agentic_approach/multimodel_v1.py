import os
import pandas as pd
import instructor
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from dotenv import load_dotenv

# Load credentials from .env file
load_dotenv()

# --- 1. THE DYNAMIC SCHEMA ---
# This is the "Master Template". Add or remove fields here; 
# the rest of the code adjusts automatically.
class LineItem(BaseModel):
    description: Optional[str] = Field(None, description="Exact text of the item description")
    qty: Optional[float] = Field(None, description="Quantity")
    uom: Optional[str] = Field(None, description="Unit of measure (e.g., kg, kwh)")
    amount: Optional[float] = Field(None, description="Line total amount")
    charge_type: Optional[str] = Field(None, description="Category: e.g., Freight, Rent, Energy")

class UnifiedInvoice(BaseModel):
    invoice_number: Optional[str] = Field(None, description="The unique ID of the invoice")
    vendor_name: Optional[str] = Field(None, description="Name of the issuing company")
    line_items: List[LineItem]

# --- 2. THE CONFIDENCE RESOLVER ---
# Maps LLM strings back to Azure's Word-Level OCR confidence
class ConfidenceMapper:
    def __init__(self, analyze_result):
        self.word_map = {}
        if analyze_result.pages:
            for page in analyze_result.pages:
                if page.words:
                    for word in page.words:
                        # Clean word for better matching
                        clean_text = word.content.strip().lower()
                        self.word_map[clean_text] = word.confidence

    def get_phrase_confidence(self, value):
        if value is None or str(value).strip() == "":
            return 0.0
        
        search_vals = str(value).strip().lower().split()
        confidences = [self.word_map.get(w, 0.5) for w in search_vals]
        
        return round(sum(confidences) / len(confidences), 2) if confidences else 0.5

# --- 3. THE EXTRACTION ENGINE ---
def process_invoice(file_path: str):
    # A. Initialize Azure Clients
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

    # B. Step 1: Azure Layout (OCR to Markdown)
    print(f"Status: Analyzing layout for {os.path.basename(file_path)}...")
    with open(file_path, "rb") as f:
        poller = di_client.begin_analyze_document(
            "prebuilt-layout",
            AnalyzeDocumentRequest(bytes_source=f.read()),
            output_content_format="markdown" # Stable string literal
        )
    result = poller.result()
    conf_resolver = ConfidenceMapper(result)

    # C. Step 2: Instructor Extraction (LLM Reasoning)
    print("Status: Extracting structured data via LLM...")
    extracted_data = ai_client.chat.completions.create(
        model="gpt-4o-mini", # Deployment name
        response_model=UnifiedInvoice,
        max_retries=3,
        messages=[
            {
                "role": "system",
                "content": (
                    "Strictly extract visible text into the schema. "
                    "Do NOT calculate totals. Do NOT fill in missing data. "
                    "Return NULL for any field not explicitly present in the text."
                )
            },
            {"role": "user", "content": result.content}
        ]
    )

    # D. Step 3: Dynamic DataFrame Creation
    all_rows = []
    for item in extracted_data.line_items:
        row_data = {}
        # Dynamic loop through Pydantic fields to prevent hardcoding
        for field, value in item.model_dump().items():
            row_data[field] = value
            # Calculate the ACTUAL OCR confidence for this specific value
            row_data[f"{field}_conf"] = conf_resolver.get_phrase_confidence(value)
        all_rows.append(row_data)

    return pd.DataFrame(all_rows), extracted_data

# --- 4. EXECUTION ---
if __name__ == "__main__":
    # Path to your test file
    FILE_PATH = "invoice_sample.pdf" 
    
    try:
        df, raw_obj = process_invoice(FILE_PATH)
        
        print("\n--- EXTRACTION RESULTS ---")
        print(f"Vendor: {raw_obj.vendor_name}")
        print(f"Invoice #: {raw_obj.invoice_number}")
        print("-" * 30)
        
        # Display the DataFrame
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', 1000)
        print(df)
        
    except Exception as e:
        print(f"Critical Error: {e}")
