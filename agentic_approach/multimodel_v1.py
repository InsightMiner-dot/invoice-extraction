import os
import pandas as pd
import instructor
from openai import AzureOpenAI
from pydantic import BaseModel, Field
from typing import List, Optional, Dict
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest, ContentFormat
from dotenv import load_dotenv

load_dotenv()

# --- 1. Dynamic Pydantic Schema ---
class LineItem(BaseModel):
    description: Optional[str] = Field(None, description="Exact text of the item description")
    qty: Optional[float] = Field(None, description="Quantity as written")
    uom: Optional[str] = Field(None, description="Unit of measure")
    amount: Optional[float] = Field(None, description="Line total")
    charge_type: Optional[str] = Field(None, description="Transport/Rental specific category")

class UnifiedInvoice(BaseModel):
    invoice_number: Optional[str] = None
    vendor_name: Optional[str] = None
    line_items: List[LineItem]

# --- 2. The Confidence Resolver Engine ---
class ConfidenceMapper:
    """Matches LLM output back to Azure OCR Confidence scores."""
    def __init__(self, analyze_result):
        self.cell_map = {}
        # We build a dictionary of {text: confidence} from the raw Azure tables
        for table in analyze_result.tables:
            for cell in table.cells:
                # Clean text to ensure better matching
                clean_text = cell.content.strip().lower()
                self.cell_map[clean_text] = cell.confidence

    def get_conf(self, value):
        if value is None: return 0.0
        search_val = str(value).strip().lower()
        # Returns the actual Azure confidence, or 0.5 if it's a partial match/hallucination
        return self.cell_map.get(search_val, 0.5)

# --- 3. Main Extraction Logic ---
def run_dynamic_extraction(file_path: str):
    # A. Azure Document Intelligence (Layout)
    di_client = DocumentIntelligenceClient(
        endpoint=os.getenv("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT"), 
        credential=AzureKeyCredential(os.getenv("AZURE_DOCUMENT_INTELLIGENCE_KEY"))
    )
    
    with open(file_path, "rb") as f:
        poller = di_client.begin_analyze_document(
            "prebuilt-layout", 
            AnalyzeDocumentRequest(bytes_source=f.read()),
            output_content_format=ContentFormat.MARKDOWN
        )
    result = poller.result()
    markdown_content = result.content
    conf_resolver = ConfidenceMapper(result)

    # B. Instructor + Azure OpenAI
    ai_client = instructor.from_openai(
        AzureOpenAI(
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            api_version="2024-02-15-preview"
        )
    )

    # STRICT INSTRUCTIONS to avoid LLM "filling in" data
    extracted_data = ai_client.chat.completions.create(
        model="gpt-4o-mini",
        response_model=UnifiedInvoice,
        messages=[
            {
                "role": "system", 
                "content": (
                    "You are a strict data extraction engine. "
                    "1. Extract ONLY information visible in the text. "
                    "2. If a field is not explicitly stated, return NULL. "
                    "3. DO NOT guess, calculate, or infer missing values. "
                    "4. If a table row is incomplete, leave the missing columns as NULL."
                )
            },
            {"role": "user", "content": markdown_content}
        ]
    )

    # C. Dynamic DataFrame Generation with Real Confidence
    rows = []
    for item in extracted_data.line_items:
        row_dict = {}
        # Dynamically iterate through Pydantic fields (no hard-coding names)
        for field_name, value in item.model_dump().items():
            row_dict[field_name] = value
            # Add the ACTUAL Azure confidence score for this specific cell
            row_dict[f"{field_name}_conf"] = conf_resolver.get_conf(value)
        rows.append(row_dict)

    return pd.DataFrame(rows)

# --- 4. Execution ---
if __name__ == "__main__":
    # Ensure this file exists in your directory
    FILE_TO_PROCESS = "test_invoice.pdf" 
    
    try:
        final_df = run_dynamic_extraction(FILE_TO_PROCESS)
        
        # Displaying the result
        print("\n--- DYNAMIC EXTRACTION WITH REAL CONFIDENCE ---")
        # Format for better visibility in console
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', 1000)
        print(final_df)
        
        # Export for your UI
        # final_df.to_csv("extraction_results.csv", index=False)
        
    except Exception as e:
        print(f"Error occurred: {e}")
