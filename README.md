# Invoice Extraction Pipeline
## Azure Document Intelligence + Azure OpenAI gpt-4o-mini

---

## Folder Structure

```
invoice_extractor/
├── pipeline.py           ← Main pipeline (do not edit for config changes)
├── fields_config.json    ← ALL field definitions, aliases, CSV columns (edit here)
├── .env                  ← Your Azure credentials (copy from .env.template)
├── .env.template         ← Credential template
├── requirements.txt
├── invoices/             ← Drop PDF invoices here
└── output/
    ├── invoices_extracted.csv   ← Final output
    └── needs_review/            ← Low-confidence extractions (JSON) for human check
```

---

## Setup

```bash
pip install -r requirements.txt
cp .env.template .env
# Fill in your Azure credentials in .env
```

---

## .env Reference

| Variable | Description |
|---|---|
| `AZURE_DI_ENDPOINT` | Azure Document Intelligence endpoint URL |
| `AZURE_DI_KEY` | Azure DI API key |
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI endpoint URL |
| `AZURE_OPENAI_KEY` | Azure OpenAI API key |
| `AZURE_OPENAI_DEPLOYMENT` | Deployment name (default: `gpt-4o-mini`) |
| `AZURE_OPENAI_API_VERSION` | API version (default: `2024-08-01-preview`) |
| `CONFIDENCE_THRESHOLD` | Score below which invoice goes to review (default: `0.80`) |
| `INPUT_FOLDER` | Folder to scan for PDFs (default: `./invoices`) |
| `OUTPUT_CSV` | Output CSV file path |
| `FIELDS_CONFIG` | Path to fields_config.json |
| `REVIEW_FOLDER` | Folder for low-confidence JSON outputs |

---

## Run

```bash
python pipeline.py
```

---

## CSV Output Columns

```
File Name | Supplier Name | Invoice Number | Invoice Date | Supplier Address |
Remit To | Ship To | Shipper Address | Bill To (with Address) | Origin |
Destination | Total Miles | Total Pieces | Total Invoice Amount | Currency |
Description | Volume/Quantity | UOM | Cost/Rate | Amount | Confidence
```

One row per line item. Header fields repeat on every row for that invoice.
Tax, GST, CGST, SGST, IGST lines each appear as their own row.

---

## Adding a New Field (no code changes needed)

Open `fields_config.json` and add a block:

```json
"po_number": {
  "csv_column": "PO Number",
  "aliases": ["po number", "purchase order", "po #", "order number"],
  "required": false
}
```

- Add to `header_fields` for a once-per-invoice field
- Add to `line_item_fields` for a per-row field
- Run pipeline — new column appears in CSV automatically

---

## Adding a New Tax Type

Open `fields_config.json` → `tax_line_items.known_tax_labels` and add the label:

```json
"known_tax_labels": [
  "gst", "cgst", "sgst", "igst", "vat", "your_new_tax_here"
]
```

---

## How Strict Grounding Works

The LLM is given `temperature=0` and explicit rules:
- **Never infer** a value not present on the document
- **Never calculate** (e.g., if subtotal + tax isn't shown, total stays null)
- **Never merge** fields (bill-to and ship-to stay separate even if similar)
- Missing fields return `null`, not a guess

Low-confidence invoices are saved as JSON in `output/needs_review/` for human correction. 
Corrections can be used as few-shot examples to improve future extractions.

---

## Confidence Score

| Score | Meaning |
|---|---|
| 1.0 | All required fields found, math checks out |
| 0.8–0.99 | Minor issues (one missing optional field or small amount variance) |
| < 0.8 | Routed to `needs_review/` folder for human check |
