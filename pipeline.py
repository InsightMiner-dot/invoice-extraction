"""
Invoice Extraction Pipeline
Azure Document Intelligence (Layout) + Azure OpenAI gpt-4o-mini
Config-driven: all fields, aliases, CSV columns defined in fields_config.json
"""

import os
import json
import csv
import logging
import re
from pathlib import Path
from datetime import datetime
from typing import Any

from dotenv import load_dotenv
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from openai import AzureOpenAI

load_dotenv()

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# 1. CONFIG LOADER
# ══════════════════════════════════════════════════════════════════════════════

def load_config(config_path: str) -> dict:
    with open(config_path, "r") as f:
        return json.load(f)


def build_csv_columns(config: dict) -> list[str]:
    """
    Derive the ordered CSV column list from fields_config.json.
    Header fields come first (in definition order), then line-item fields.
    """
    header_cols = [
        v["csv_column"]
        for k, v in config["header_fields"].items()
    ]
    line_item_cols = [
        v["csv_column"]
        for k, v in config["line_item_fields"].items()
    ]
    return header_cols + line_item_cols


def build_extraction_schema(config: dict) -> dict:
    """
    Build the JSON schema we ask the LLM to return.
    Derived entirely from fields_config.json — no hardcoding.
    """
    header_schema = {}
    for field_key, meta in config["header_fields"].items():
        if meta.get("source") == "system":
            continue
        header_schema[field_key] = {
            "type": "string or null",
            "aliases_to_look_for": meta["aliases"],
            "csv_column": meta["csv_column"]
        }

    line_item_schema = {}
    for field_key, meta in config["line_item_fields"].items():
        line_item_schema[field_key] = {
            "type": "string or null",
            "aliases_to_look_for": meta["aliases"],
            "csv_column": meta["csv_column"]
        }

    tax_labels = config["tax_line_items"]["known_tax_labels"]

    return {
        "header": header_schema,
        "line_items": {
            "_description": "Array of ALL line items including charges. Each tax/GST/CGST/SGST/IGST/surcharge found on the document MUST be a separate line item row.",
            "_tax_labels_to_watch": tax_labels,
            "fields": line_item_schema
        }
    }


# ══════════════════════════════════════════════════════════════════════════════
# 2. AZURE DOCUMENT INTELLIGENCE — raw layout extraction
# ══════════════════════════════════════════════════════════════════════════════

def extract_raw_layout(pdf_path: str) -> str:
    """
    Use Azure DI prebuilt-layout to extract raw text + tables.
    Returns a plain-text representation with tables preserved as pipe-delimited rows.
    No schema applied here — LLM handles semantics.
    """
    client = DocumentAnalysisClient(
        endpoint=os.getenv("AZURE_DI_ENDPOINT"),
        credential=AzureKeyCredential(os.getenv("AZURE_DI_KEY"))
    )

    with open(pdf_path, "rb") as f:
        poller = client.begin_analyze_document("prebuilt-layout", f)
    result = poller.result()

    sections = []

    # Raw text lines (preserve reading order)
    page_lines = []
    for page in result.pages:
        for line in page.lines:
            page_lines.append(line.content)
    if page_lines:
        sections.append("=== DOCUMENT TEXT ===")
        sections.extend(page_lines)

    # Tables (pipe-delimited for LLM readability)
    for t_idx, table in enumerate(result.tables):
        rows: dict[int, dict[int, str]] = {}
        for cell in table.cells:
            rows.setdefault(cell.row_index, {})[cell.column_index] = cell.content
        sections.append(f"\n=== TABLE {t_idx + 1} ===")
        for r in sorted(rows):
            row_text = " | ".join(
                rows[r].get(c, "").strip()
                for c in sorted(rows[r])
            )
            sections.append(row_text)

    return "\n".join(sections)


# ══════════════════════════════════════════════════════════════════════════════
# 3. AZURE OPENAI — semantic extraction with strict grounding
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are a strict invoice data extraction engine.

RULES — follow exactly:
1. ONLY extract values that are explicitly present as text on the document.
2. NEVER infer, calculate, guess, or assume any value.
3. NEVER combine or merge fields unless they appear together on the document.
4. If a field is not found, return null — do not guess or leave a blank.
5. For line items: extract EVERY row from the document as a separate object, preserving exact text.
6. TAX RULE: Any line containing tax, GST, CGST, SGST, IGST, VAT, surcharge, cess, duty, or any charge that is NOT a freight/service line — extract it as a separate line_item object with its exact label as the description.
7. Return ONLY valid JSON. No explanation, no markdown, no preamble.
8. Preserve exact values as written on the document (do not reformat dates, amounts, or addresses).
"""


def build_user_prompt(raw_text: str, extraction_schema: dict) -> str:
    schema_str = json.dumps(extraction_schema, indent=2)
    return f"""Extract all invoice data from the document text below.

Return a JSON object with this exact structure:
{schema_str}

IMPORTANT:
- Use null (not empty string) for any field not found on the document.
- For line_items: include every line row as a separate object, including all tax/GST/CGST/SGST/IGST/surcharge lines as individual line_item entries.
- Do NOT merge tax into a header field. Tax lines go into line_items only.
- Do NOT infer values. Only extract what is literally on the document.

=== DOCUMENT TEXT ===
{raw_text}
"""


def extract_with_llm(raw_text: str, extraction_schema: dict) -> dict:
    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_KEY"),
        api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview")
    )

    response = client.chat.completions.create(
        model=os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini"),
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_prompt(raw_text, extraction_schema)}
        ],
        temperature=0,          # Deterministic — no creativity
        max_tokens=4096,
        response_format={"type": "json_object"}   # Forces valid JSON output
    )

    raw_json = response.choices[0].message.content
    return json.loads(raw_json)


# ══════════════════════════════════════════════════════════════════════════════
# 4. VALIDATION + CONFIDENCE SCORING
# ══════════════════════════════════════════════════════════════════════════════

def validate_and_score(extracted: dict, config: dict, file_name: str) -> tuple[dict, float, list[str]]:
    """
    Check required fields and basic math integrity.
    Returns (enriched_extracted, confidence_score, list_of_issues).
    """
    issues = []
    checks = 0

    header = extracted.get("header", {})
    line_items = extracted.get("line_items", [])

    # Required field checks
    for field_key, meta in config["header_fields"].items():
        if meta.get("source") == "system":
            continue
        if meta.get("required", False):
            checks += 1
            val = header.get(field_key)
            if not val or str(val).strip().lower() in ("null", "none", ""):
                issues.append(f"Required field missing: {meta['csv_column']}")

    # Line items present
    checks += 1
    if not line_items:
        issues.append("No line items extracted")

    # Math check: sum of line item amounts vs total (if both present)
    try:
        total_str = header.get("total_invoice_amount")
        if total_str and line_items:
            checks += 1
            total_val = float(re.sub(r"[^\d.]", "", str(total_str)))
            line_sum = sum(
                float(re.sub(r"[^\d.]", "", str(li.get("amount", "") or "")))
                for li in line_items
                if li.get("amount") and str(li.get("amount")).strip() not in ("", "null", "None")
            )
            if line_sum > 0 and abs(line_sum - total_val) / max(total_val, 1) > 0.05:
                issues.append(
                    f"Amount mismatch: line items sum {line_sum:.2f} vs total {total_val:.2f}"
                )
    except (ValueError, TypeError):
        pass  # Non-numeric totals — skip math check

    confidence = round(1 - (len(issues) / max(checks, 1)), 2)

    # Inject system fields
    header["file_name"] = file_name

    return extracted, confidence, issues


# ══════════════════════════════════════════════════════════════════════════════
# 5. CSV WRITER
# ══════════════════════════════════════════════════════════════════════════════

def flatten_to_csv_rows(extracted: dict, config: dict, confidence: float) -> list[dict]:
    """
    Convert extracted JSON → list of CSV row dicts.
    One row per line item. Header fields repeat on every row.
    """
    header = extracted.get("header", {})
    line_items = extracted.get("line_items", [])

    # Build header value map: csv_column → value
    header_row = {}
    for field_key, meta in config["header_fields"].items():
        col = meta["csv_column"]
        val = header.get(field_key)
        header_row[col] = "" if val is None or str(val).lower() in ("null", "none") else str(val).strip()

    header_row["Confidence"] = str(confidence)

    # If no line items, write one row with empty line-item columns
    if not line_items:
        row = dict(header_row)
        for field_key, meta in config["line_item_fields"].items():
            row[meta["csv_column"]] = ""
        return [row]

    rows = []
    for li in line_items:
        row = dict(header_row)
        for field_key, meta in config["line_item_fields"].items():
            val = li.get(field_key)
            row[meta["csv_column"]] = "" if val is None or str(val).lower() in ("null", "none") else str(val).strip()
        rows.append(row)

    return rows


def write_csv(all_rows: list[dict], csv_columns: list[str], output_path: str):
    """Append rows to the output CSV, writing header only once."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    write_header = not output_path.exists()
    all_columns = csv_columns + ["Confidence"]

    with open(output_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_columns, extrasaction="ignore")
        if write_header:
            writer.writeheader()
        writer.writerows(all_rows)


# ══════════════════════════════════════════════════════════════════════════════
# 6. MAIN PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

def process_invoice(pdf_path: str, config: dict, extraction_schema: dict, csv_columns: list[str]) -> dict:
    file_name = Path(pdf_path).name
    log.info(f"Processing: {file_name}")

    result = {
        "file": file_name,
        "status": "success",
        "confidence": None,
        "issues": [],
        "rows_written": 0
    }

    try:
        # Step 1: Azure DI — raw layout
        log.info(f"  [1/4] Azure DI layout extraction...")
        raw_text = extract_raw_layout(pdf_path)

        # Step 2: Azure OpenAI — semantic extraction
        log.info(f"  [2/4] Azure OpenAI gpt-4o-mini extraction...")
        extracted = extract_with_llm(raw_text, extraction_schema)

        # Step 3: Validate + score
        log.info(f"  [3/4] Validating...")
        extracted, confidence, issues = validate_and_score(extracted, config, file_name)
        result["confidence"] = confidence
        result["issues"] = issues

        if issues:
            log.warning(f"  Issues: {issues}")

        # Step 4: Write CSV
        log.info(f"  [4/4] Writing CSV...")
        rows = flatten_to_csv_rows(extracted, config, confidence)
        output_csv = os.getenv("OUTPUT_CSV", "./output/invoices_extracted.csv")
        write_csv(rows, csv_columns, output_csv)
        result["rows_written"] = len(rows)

        # Save to review folder if below confidence threshold
        threshold = float(os.getenv("CONFIDENCE_THRESHOLD", 0.80))
        if confidence < threshold:
            review_dir = Path(os.getenv("REVIEW_FOLDER", "./output/needs_review"))
            review_dir.mkdir(parents=True, exist_ok=True)
            review_json = review_dir / f"{Path(pdf_path).stem}_review.json"
            with open(review_json, "w") as f:
                json.dump({
                    "file": file_name,
                    "confidence": confidence,
                    "issues": issues,
                    "extracted": extracted
                }, f, indent=2)
            result["status"] = "needs_review"
            log.warning(f"  Low confidence ({confidence}). Saved for review: {review_json}")

        log.info(f"  Done. Confidence={confidence}, Rows={len(rows)}")

    except Exception as e:
        log.error(f"  FAILED: {e}", exc_info=True)
        result["status"] = "error"
        result["error"] = str(e)

    return result


def run_pipeline():
    config_path = os.getenv("FIELDS_CONFIG", "./fields_config.json")
    input_folder = os.getenv("INPUT_FOLDER", "./invoices")

    log.info("=" * 60)
    log.info("Invoice Extraction Pipeline — Azure DI + gpt-4o-mini")
    log.info(f"Config     : {config_path}")
    log.info(f"Input      : {input_folder}")
    log.info(f"Output CSV : {os.getenv('OUTPUT_CSV')}")
    log.info("=" * 60)

    # Load config — everything derives from this
    config = load_config(config_path)
    csv_columns = build_csv_columns(config)
    extraction_schema = build_extraction_schema(config)

    # Find all PDFs
    pdf_files = list(Path(input_folder).glob("**/*.pdf"))
    if not pdf_files:
        log.warning(f"No PDF files found in {input_folder}")
        return

    log.info(f"Found {len(pdf_files)} PDF(s) to process")

    summary = {"success": 0, "needs_review": 0, "error": 0}

    for pdf_path in pdf_files:
        result = process_invoice(str(pdf_path), config, extraction_schema, csv_columns)
        summary[result.get("status", "error")] = summary.get(result.get("status", "error"), 0) + 1

    log.info("=" * 60)
    log.info(f"COMPLETE — Success: {summary['success']} | Review: {summary['needs_review']} | Errors: {summary['error']}")
    log.info(f"Output CSV: {os.getenv('OUTPUT_CSV')}")
    log.info("=" * 60)


if __name__ == "__main__":
    run_pipeline()
