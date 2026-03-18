"""
Invoice Extraction Pipeline — Streamlit UI
Azure Document Intelligence + Azure OpenAI gpt-4o-mini

Batch processing features:
  - Concurrent PDF processing via ThreadPoolExecutor
  - Per-file isolation: one file failing never stops the batch
  - Retry logic with exponential back-off on Azure throttling (429)
  - Deduplication guard: re-uploading same filename replaces, not duplicates
  - Two-level progress: overall batch bar + per-file status ticker
  - Batch summary table at the end
  - Page(s) column per line item
  - Multi-invoice-per-PDF detection
"""

import os
import json
import csv
import io
import re
import time
import hashlib
import concurrent.futures
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from typing import Callable

import streamlit as st
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Invoice Intelligence",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

[data-testid="stSidebar"] { background:#0f1117; border-right:1px solid #1e2330; }
[data-testid="stSidebar"] * { color:#c9d1e0 !important; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color:#fff !important; }

.main .block-container { padding-top:2rem; padding-bottom:3rem; max-width:1400px; }

.header-banner {
    background:linear-gradient(135deg,#0f1117 0%,#1a1f2e 50%,#0d1117 100%);
    border:1px solid #1e2330; border-radius:16px;
    padding:2rem 2.5rem; margin-bottom:2rem;
    display:flex; align-items:center; gap:1.5rem;
}
.header-title  { font-size:2rem; font-weight:600; color:#fff; margin:0; letter-spacing:-0.5px; }
.header-subtitle { color:#6b7898; font-size:0.9rem; margin:.25rem 0 0 0; font-family:'DM Mono',monospace; }

.metric-card { background:#0f1117; border:1px solid #1e2330; border-radius:12px; padding:1.25rem 1.5rem; text-align:center; }
.metric-value { font-size:2rem; font-weight:600; color:#4f9cf9; line-height:1; }
.metric-label { font-size:.75rem; color:#6b7898; text-transform:uppercase; letter-spacing:1px; margin-top:.4rem; font-family:'DM Mono',monospace; }

.pill-success { background:#0d2818; color:#4ade80; border:1px solid #166534; padding:3px 12px; border-radius:20px; font-size:.75rem; font-family:'DM Mono',monospace; display:inline-block; }
.pill-warning { background:#2d1f00; color:#fbbf24; border:1px solid #854d0e; padding:3px 12px; border-radius:20px; font-size:.75rem; font-family:'DM Mono',monospace; display:inline-block; }
.pill-error   { background:#2d0f0f; color:#f87171; border:1px solid #991b1b; padding:3px 12px; border-radius:20px; font-size:.75rem; font-family:'DM Mono',monospace; display:inline-block; }
.pill-running { background:#0a1f3d; color:#60a5fa; border:1px solid #1d4ed8; padding:3px 12px; border-radius:20px; font-size:.75rem; font-family:'DM Mono',monospace; display:inline-block; }

.batch-row {
    display:grid; grid-template-columns:2fr 1fr 1fr 80px 120px;
    align-items:center; gap:.75rem;
    padding:.6rem 1rem; border-radius:8px;
    border:1px solid #1e2330; margin-bottom:.4rem;
    font-size:.82rem;
}
.batch-row:hover { border-color:#2d3a52; }
.file-name { font-family:'DM Mono',monospace; font-size:.82rem; color:#c9d1e0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.conf-bar-wrap { width:80px; height:5px; background:#1e2330; border-radius:3px; overflow:hidden; display:inline-block; vertical-align:middle; margin-right:6px; }
.conf-bar-fill { height:100%; border-radius:3px; }
.section-label { font-family:'DM Mono',monospace; font-size:.7rem; text-transform:uppercase; letter-spacing:2px; color:#6b7898; margin-bottom:.75rem; }
.multi-banner { background:#1c1400; border:1px solid #854d0e; border-radius:8px; padding:.75rem 1rem; font-family:'DM Mono',monospace; font-size:.8rem; color:#fbbf24; }
.page-badge { background:#0a1f3d; color:#60a5fa; border:1px solid #1d4ed8; padding:1px 7px; border-radius:4px; font-size:.72rem; font-family:'DM Mono',monospace; display:inline-block; margin:1px; }

.stButton > button { background:#4f9cf9; color:#fff; border:none; border-radius:8px; font-family:'DM Sans',sans-serif; font-weight:500; padding:.6rem 1.5rem; width:100%; transition:background .2s; }
.stButton > button:hover { background:#3b8de8; }
.stButton > button:disabled { background:#1e2330; color:#6b7898; }
div[data-testid="stExpander"] { border:1px solid #1e2330; border-radius:10px; background:#0f1117; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CONFIG HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def load_config(path: str) -> dict:
    with open(path) as f:
        return json.load(f)


def build_csv_columns(config: dict) -> list[str]:
    header_cols   = [v["csv_column"] for v in config["header_fields"].values()]
    li_cols       = [v["csv_column"] for v in config["line_item_fields"].values()]
    file_col      = header_cols[:1]        # File Name always first
    rest_header   = header_cols[1:]
    return file_col + ["Page(s)", "Invoice Split #"] + rest_header + li_cols + ["Confidence"]


def build_extraction_schema(config: dict) -> dict:
    header_schema = {}
    for key, meta in config["header_fields"].items():
        if meta.get("source") == "system":
            continue
        header_schema[key] = {
            "type": "string or null",
            "aliases_to_look_for": meta["aliases"],
            "csv_column": meta["csv_column"]
        }
    li_schema = {}
    for key, meta in config["line_item_fields"].items():
        li_schema[key] = {
            "type": "string or null",
            "aliases_to_look_for": meta["aliases"],
            "csv_column": meta["csv_column"]
        }
    return {
        "header": header_schema,
        "line_items": {
            "_description": (
                "Array of ALL charge rows. Every tax/GST/CGST/SGST/IGST/surcharge/cess/duty "
                "MUST be its own separate line_item object. "
                "Set page_number (int, 1-based) on every line item."
            ),
            "_tax_labels": config["tax_line_items"]["known_tax_labels"],
            "fields": {
                **li_schema,
                "page_number": {
                    "type": "integer",
                    "description": "1-based page number where this row appears"
                }
            }
        }
    }


def file_hash(pdf_bytes: bytes) -> str:
    """MD5 of file content — used to deduplicate re-uploads."""
    return hashlib.md5(pdf_bytes).hexdigest()


# ══════════════════════════════════════════════════════════════════════════════
# AZURE DI — page-aware layout
# ══════════════════════════════════════════════════════════════════════════════

def extract_layout_by_page(pdf_bytes: bytes) -> tuple[list[dict], int]:
    from azure.ai.formrecognizer import DocumentAnalysisClient
    from azure.core.credentials import AzureKeyCredential

    client = DocumentAnalysisClient(
        endpoint=os.getenv("AZURE_DI_ENDPOINT"),
        credential=AzureKeyCredential(os.getenv("AZURE_DI_KEY"))
    )
    poller = client.begin_analyze_document("prebuilt-layout", pdf_bytes)
    result = poller.result()
    total_pages = len(result.pages)

    page_lines: dict[int, list] = defaultdict(list)
    for page in result.pages:
        for line in page.lines:
            page_lines[page.page_number].append(line.content)

    page_tables: dict[int, list] = defaultdict(list)
    for t_idx, table in enumerate(result.tables):
        rows: dict[int, dict] = {}
        first_page = 1
        for cell in table.cells:
            if cell.bounding_regions:
                first_page = cell.bounding_regions[0].page_number
            rows.setdefault(cell.row_index, {})[cell.column_index] = cell.content
        t_lines = [
            " | ".join(rows[r].get(c, "").strip() for c in sorted(rows[r]))
            for r in sorted(rows)
        ]
        page_tables[first_page].append((t_idx + 1, t_lines))

    pages_data = []
    for pnum in sorted(page_lines.keys()):
        parts = [f"=== PAGE {pnum} ==="]
        parts.extend(page_lines[pnum])
        for t_num, t_lines in page_tables.get(pnum, []):
            parts.append(f"--- Table {t_num} ---")
            parts.extend(t_lines)
        pages_data.append({"page_num": pnum, "text": "\n".join(parts)})

    return pages_data, total_pages


# ══════════════════════════════════════════════════════════════════════════════
# MULTI-INVOICE DETECTION
# ══════════════════════════════════════════════════════════════════════════════

INV_PATTERNS = [
    r'\b(?:invoice|inv|bill)\s*(?:no|num|number|#|:)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-\/]{2,20})\b',
    r'\b([A-Z]{2,5}[-\/]?\d{4,12})\b',
    r'\b(INV[-\/]?\d{3,12})\b',
]

def detect_invoice_per_page(pages_data: list[dict]) -> dict[int, str | None]:
    result = {}
    for pd_item in pages_data:
        found = None
        for pat in INV_PATTERNS:
            m = re.search(pat, pd_item["text"], re.IGNORECASE)
            if m:
                found = m.group(1).upper().strip()
                break
        result[pd_item["page_num"]] = found
    return result


def group_pages_into_segments(page_inv_map: dict[int, str | None]) -> list[dict]:
    """
    - Same inv number → extend current segment
    - Different inv number → new segment
    - No inv number → continuation of current segment
    """
    groups: list[dict] = []
    curr_inv: str | None = None
    curr_pages: list[int] = []

    for pnum in sorted(page_inv_map.keys()):
        detected = page_inv_map[pnum]
        if detected is None:
            curr_pages.append(pnum)
        elif detected == curr_inv:
            curr_pages.append(pnum)
        else:
            if curr_pages:
                groups.append({"invoice_number": curr_inv, "pages": curr_pages})
            curr_inv   = detected
            curr_pages = [pnum]

    if curr_pages:
        groups.append({"invoice_number": curr_inv, "pages": curr_pages})

    return groups


# ══════════════════════════════════════════════════════════════════════════════
# LLM EXTRACTION WITH RETRY
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are a strict invoice data extraction engine.

ABSOLUTE RULES:
1. ONLY extract values explicitly present as text in the document.
2. NEVER infer, calculate, guess, or assume any value.
3. Return null for any field not found — never guess or use empty string.
4. Extract EVERY line row as a separate object in line_items.
5. TAX RULE: Every line containing tax / GST / CGST / SGST / IGST / VAT / HST /
   surcharge / cess / duty → its own separate line_item with exact label as description.
6. PAGE RULE: Set page_number (int, 1-based) on every line_item.
7. Return ONLY valid JSON. No markdown, no explanation."""


def extract_with_llm(
    segment_text: str,
    schema: dict,
    page_hint: str,
    max_retries: int = 3
) -> dict:
    from openai import AzureOpenAI, RateLimitError, APIError

    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_KEY"),
        api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview")
    )

    user_msg = f"""Extract all invoice data from the document pages below.
Segment: {page_hint}

Return JSON matching this structure exactly:
{json.dumps(schema, indent=2)}

Rules:
- null for any missing field (not empty string).
- Tax/GST/CGST/SGST/IGST/surcharge → individual line_item rows only.
- page_number on every line item.
- Do NOT infer. Extract only text present on the document.

=== DOCUMENT PAGES ===
{segment_text}"""

    for attempt in range(1, max_retries + 1):
        try:
            resp = client.chat.completions.create(
                model=os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini"),
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_msg}
                ],
                temperature=0,
                max_tokens=4096,
                response_format={"type": "json_object"}
            )
            return json.loads(resp.choices[0].message.content)

        except RateLimitError:
            wait = 2 ** attempt          # 2s, 4s, 8s
            if attempt < max_retries:
                time.sleep(wait)
            else:
                raise

        except (json.JSONDecodeError, APIError) as e:
            if attempt < max_retries:
                time.sleep(1)
            else:
                raise


# ══════════════════════════════════════════════════════════════════════════════
# VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

def validate_and_score(
    extracted: dict, config: dict,
    file_name: str, pages: list[int], split_num: int
) -> tuple[dict, float, list[str]]:

    issues = []
    checks = 0
    header     = extracted.get("header", {})
    line_items = extracted.get("line_items", [])

    for key, meta in config["header_fields"].items():
        if meta.get("source") == "system":
            continue
        if meta.get("required", False):
            checks += 1
            val = header.get(key)
            if not val or str(val).strip().lower() in ("null", "none", ""):
                issues.append(f"Missing required: {meta['csv_column']}")

    checks += 1
    if not line_items:
        issues.append("No line items extracted")

    try:
        total_str = header.get("total_invoice_amount")
        if total_str and line_items:
            checks += 1
            total_val = float(re.sub(r"[^\d.]", "", str(total_str)))
            line_sum  = sum(
                float(re.sub(r"[^\d.]", "", str(li.get("amount") or "")))
                for li in line_items
                if li.get("amount") and str(li.get("amount")).strip() not in ("", "null", "None")
            )
            if line_sum > 0 and abs(line_sum - total_val) / max(total_val, 1) > 0.05:
                issues.append(f"Amount mismatch: lines={line_sum:.2f} total={total_val:.2f}")
    except (ValueError, TypeError):
        pass

    confidence = round(1 - (len(issues) / max(checks, 1)), 2)
    header["file_name"] = file_name
    header["pages"]     = ",".join(str(p) for p in sorted(pages))
    header["split_num"] = str(split_num)
    return extracted, confidence, issues


# ══════════════════════════════════════════════════════════════════════════════
# FLATTEN → CSV ROWS
# ══════════════════════════════════════════════════════════════════════════════

def flatten_to_rows(extracted: dict, config: dict, confidence: float) -> list[dict]:
    header     = extracted.get("header", {})
    line_items = extracted.get("line_items", [])
    inv_pages  = header.get("pages", "")
    split_num  = header.get("split_num", "1")

    base = {}
    for key, meta in config["header_fields"].items():
        col = meta["csv_column"]
        val = header.get(key)
        base[col] = "" if val is None or str(val).lower() in ("null", "none") else str(val).strip()

    base["Page(s)"]         = inv_pages
    base["Invoice Split #"] = split_num
    base["Confidence"]      = str(confidence)

    if not line_items:
        row = dict(base)
        for meta in config["line_item_fields"].values():
            row[meta["csv_column"]] = ""
        return [row]

    rows = []
    for li in line_items:
        row = dict(base)
        li_page = li.get("page_number")
        if li_page and str(li_page).strip() not in ("", "null", "None", "0"):
            row["Page(s)"] = str(li_page)
        for key, meta in config["line_item_fields"].items():
            val = li.get(key)
            row[meta["csv_column"]] = (
                "" if val is None or str(val).lower() in ("null", "none")
                else str(val).strip()
            )
        rows.append(row)
    return rows


def rows_to_csv_bytes(rows: list, columns: list) -> bytes:
    buf = io.StringIO()
    w   = csv.DictWriter(buf, fieldnames=columns, extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue().encode("utf-8")


# ══════════════════════════════════════════════════════════════════════════════
# SINGLE PDF PROCESSOR  (runs in a thread — must be self-contained)
# ══════════════════════════════════════════════════════════════════════════════

def process_one_pdf(
    pdf_bytes: bytes,
    file_name: str,
    config: dict,
    schema: dict,
    threshold: float,
    progress_queue           # thread-safe list for status updates
) -> tuple[list[dict], list[dict]]:
    """
    Fully isolated processor for one PDF file.
    Returns (csv_rows, segment_results).
    Any exception is caught and returned as an error segment — never propagates.
    """

    def push(msg: str, icon: str = "·"):
        progress_queue.append({"file": file_name, "msg": msg, "icon": icon})

    all_rows:    list[dict] = []
    seg_results: list[dict] = []

    try:
        push("Azure DI: extracting layout…", "📄")
        pages_data, total_pages = extract_layout_by_page(pdf_bytes)
        push(f"Got {total_pages} page(s) · {sum(len(p['text']) for p in pages_data):,} chars", "✓")

        push("Detecting invoice numbers per page…", "🔍")
        page_inv_map = detect_invoice_per_page(pages_data)
        unique_invs  = {v for v in page_inv_map.values() if v}
        push(f"Invoice numbers found: {unique_invs or '(none — single invoice)'}", "✓")

        groups = group_pages_into_segments(page_inv_map)
        multi  = len(groups) > 1
        if multi:
            push(f"⚠️ {len(groups)} invoice segments in this PDF", "⚠️")

        page_text = {pd["page_num"]: pd["text"] for pd in pages_data}

        for idx, group in enumerate(groups, 1):
            seg_pages    = group["pages"]
            detected_inv = group.get("invoice_number", "—")
            page_hint    = f"pages {','.join(str(p) for p in seg_pages)} of {total_pages}"

            push(f"Segment {idx}/{len(groups)}: {page_hint} · inv# {detected_inv}", "🤖")

            seg_text = "\n\n".join(page_text[p] for p in seg_pages if p in page_text)

            try:
                extracted = extract_with_llm(seg_text, schema, page_hint)
                extracted, conf, issues = validate_and_score(
                    extracted, config, file_name, seg_pages, idx
                )
                rows = flatten_to_rows(extracted, config, conf)
                all_rows.extend(rows)

                status = "success" if conf >= threshold else "needs_review"
                push(f"Segment {idx} → {status} · conf {conf:.0%} · {len(rows)} row(s)",
                     "✅" if status == "success" else "⚠️")

                seg_results.append({
                    "file":          file_name,
                    "split_num":     idx,
                    "total_splits":  len(groups),
                    "pages":         seg_pages,
                    "detected_inv":  detected_inv,
                    "status":        status,
                    "confidence":    conf,
                    "issues":        issues,
                    "rows":          len(rows),
                    "extracted":     extracted,
                    "multi_invoice": multi,
                })

            except Exception as seg_err:
                push(f"Segment {idx} FAILED: {seg_err}", "❌")
                seg_results.append({
                    "file":         file_name,
                    "split_num":    idx,
                    "total_splits": len(groups),
                    "pages":        seg_pages,
                    "detected_inv": detected_inv,
                    "status":       "error",
                    "error":        str(seg_err),
                    "multi_invoice": multi,
                })

    except Exception as file_err:
        push(f"File-level failure: {file_err}", "❌")
        seg_results.append({
            "file":   file_name,
            "status": "error",
            "error":  str(file_err),
        })

    push(f"Done — {len(all_rows)} rows total", "🏁")
    return all_rows, seg_results


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════

DEFAULTS = {
    "results":   [],   # list of segment result dicts
    "all_rows":  [],   # list of CSV row dicts
    "file_hashes": {}, # {filename: md5} — deduplication
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    st.markdown("---")

    st.markdown('<div class="section-label">Azure Credentials</div>', unsafe_allow_html=True)
    di_ok  = bool(os.getenv("AZURE_DI_ENDPOINT") and os.getenv("AZURE_DI_KEY"))
    oai_ok = bool(os.getenv("AZURE_OPENAI_ENDPOINT") and os.getenv("AZURE_OPENAI_KEY"))

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f'<span class="{"pill-success" if di_ok else "pill-error"}">DI {"✓" if di_ok else "✗"}</span>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<span class="{"pill-success" if oai_ok else "pill-error"}">OpenAI {"✓" if oai_ok else "✗"}</span>', unsafe_allow_html=True)
    if not di_ok or not oai_ok:
        st.warning("Fill credentials in `.env` and restart.")

    st.markdown("---")
    st.markdown('<div class="section-label">Model</div>', unsafe_allow_html=True)
    st.code(os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini"), language=None)

    st.markdown("---")
    st.markdown('<div class="section-label">Batch Settings</div>', unsafe_allow_html=True)
    max_workers = st.slider("Concurrent files", 1, 5, 3,
        help="Number of PDFs processed in parallel. Keep ≤3 to avoid Azure rate limits.")
    threshold = st.slider("Confidence threshold", 0.5, 1.0,
                          float(os.getenv("CONFIDENCE_THRESHOLD", 0.80)), 0.05)

    st.markdown("---")
    st.markdown('<div class="section-label">Fields Config</div>', unsafe_allow_html=True)
    config_path = st.text_input("Path",
                                value=os.getenv("FIELDS_CONFIG", "./fields_config.json"),
                                label_visibility="collapsed")
    try:
        config = load_config(config_path)
        fc = len(config["header_fields"]) + len(config["line_item_fields"])
        st.markdown(f'<span class="pill-success">{fc} fields loaded</span>', unsafe_allow_html=True)
    except Exception as e:
        st.markdown('<span class="pill-error">Config error</span>', unsafe_allow_html=True)
        st.error(str(e))
        config = None

    st.markdown("---")
    if st.button("🗑  Clear All Results"):
        for k, v in DEFAULTS.items():
            st.session_state[k] = type(v)()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# HEADER + TABS
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="header-banner">
  <div style="font-size:3rem;line-height:1">🧾</div>
  <div>
    <div class="header-title">Invoice Intelligence</div>
    <div class="header-subtitle">Azure DI &nbsp;·&nbsp; gpt-4o-mini &nbsp;·&nbsp; Parallel batch &nbsp;·&nbsp; Page tracking &nbsp;·&nbsp; Multi-invoice detection</div>
  </div>
</div>
""", unsafe_allow_html=True)

tab_extract, tab_results, tab_config = st.tabs(
    ["📤  Extract", "📊  Results", "🔧  Field Config"]
)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 — EXTRACT
# ─────────────────────────────────────────────────────────────────────────────

with tab_extract:

    # ── Upload + summary columns ──────────────────────────────────────────────
    col_up, col_stat = st.columns([3, 2], gap="large")

    with col_up:
        st.markdown('<div class="section-label">Upload Invoice PDFs</div>', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "Drop PDFs here", type=["pdf"],
            accept_multiple_files=True, label_visibility="collapsed"
        )

        if uploaded_files:
            # Show queued files with dedup indicator
            existing_hashes = st.session_state.file_hashes
            st.markdown(f"**{len(uploaded_files)} file(s) queued**")
            for uf in uploaded_files:
                sz   = round(len(uf.getvalue()) / 1024, 1)
                fhash = file_hash(uf.getvalue())
                is_dup = uf.name in existing_hashes and existing_hashes[uf.name] == fhash
                dup_tag = ' &nbsp;<span style="color:#fbbf24;font-size:.72rem">↻ re-upload</span>' if is_dup else ""
                st.markdown(
                    f'<div class="batch-row">'
                    f'<span class="file-name">📄 {uf.name}{dup_tag}</span>'
                    f'<span style="color:#6b7898;font-family:\'DM Mono\',monospace;font-size:.8rem">{sz} KB</span>'
                    f'</div>', unsafe_allow_html=True
                )

        st.markdown("<br>", unsafe_allow_html=True)
        can_run = bool(uploaded_files and config and di_ok and oai_ok)
        run_btn = st.button("▶  Run Batch Extraction", disabled=not can_run, use_container_width=True)

        if not can_run and uploaded_files:
            if not (di_ok and oai_ok):
                st.caption("⚠️ Azure credentials missing — check `.env`")
            if not config:
                st.caption("⚠️ fields_config.json not found or invalid")

    with col_stat:
        st.markdown('<div class="section-label">Batch Summary</div>', unsafe_allow_html=True)
        results  = st.session_state.results
        success  = sum(1 for r in results if r.get("status") == "success")
        review   = sum(1 for r in results if r.get("status") == "needs_review")
        errors   = sum(1 for r in results if r.get("status") == "error")
        multi_ct = sum(1 for r in results if r.get("multi_invoice"))

        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{success}</div><div class="metric-label">Success</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color:#fbbf24">{review}</div><div class="metric-label">Review</div></div>', unsafe_allow_html=True)
        with m3:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color:#f87171">{errors}</div><div class="metric-label">Errors</div></div>', unsafe_allow_html=True)

        if results:
            unique_files_done = len({r["file"] for r in results})
            confs = [r["confidence"] for r in results if r.get("confidence")]
            total_rows = sum(r.get("rows", 0) for r in results)
            st.markdown(f"**Files processed:** `{unique_files_done}`")
            st.markdown(f"**Avg confidence:** `{sum(confs)/len(confs):.0%}`" if confs else "")
            st.markdown(f"**Total rows:** `{total_rows}`")
            if multi_ct:
                st.markdown(f'<div class="multi-banner">⚠️ {multi_ct} segment(s) from multi-invoice PDFs</div>', unsafe_allow_html=True)

    # ── BATCH RUN ────────────────────────────────────────────────────────────
    if run_btn and uploaded_files and config:
        schema      = build_extraction_schema(config)
        csv_columns = build_csv_columns(config)

        # Deduplication: if same filename+content already processed, replace its rows
        files_to_run = uploaded_files   # process all uploaded — UI shows re-upload tag

        # Shared progress queue (append-only, thread-safe via GIL on list.append)
        progress_queue: list[dict] = []

        # Overall progress bar + per-file status area
        overall_bar  = st.progress(0.0, text=f"Starting batch · {len(files_to_run)} file(s)…")
        status_area  = st.empty()
        file_status: dict[str, str] = {uf.name: "queued" for uf in files_to_run}

        def render_status():
            lines = []
            icons = {"queued": "○", "running": "●", "done": "✓", "error": "✗"}
            colors = {"queued": "#6b7898", "running": "#60a5fa", "done": "#4ade80", "error": "#f87171"}
            for fname, state in file_status.items():
                color = colors.get(state, "#6b7898")
                icon  = icons.get(state, "·")
                lines.append(
                    f'<span style="color:{color};font-family:\'DM Mono\',monospace;font-size:.8rem">'
                    f'{icon} {fname}</span>'
                )
            status_area.markdown("&nbsp;&nbsp;".join(lines), unsafe_allow_html=True)

        render_status()

        # ── Submit all files to thread pool ──────────────────────────────────
        futures: dict[concurrent.futures.Future, str] = {}
        all_rows_new  = []
        new_results   = []
        completed     = 0

        # Remove existing rows for files being re-processed (dedup)
        reprocess_names = {uf.name for uf in files_to_run}
        st.session_state.all_rows = [
            r for r in st.session_state.all_rows
            if r.get("File Name") not in reprocess_names
        ]
        st.session_state.results = [
            r for r in st.session_state.results
            if r.get("file") not in reprocess_names
        ]

        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            for uf in files_to_run:
                file_status[uf.name] = "running"
                render_status()
                fut = executor.submit(
                    process_one_pdf,
                    uf.getvalue(), uf.name,
                    config, schema, threshold,
                    progress_queue
                )
                futures[fut] = uf.name

            # Poll until all done
            while futures:
                done_futs, _ = concurrent.futures.wait(
                    list(futures.keys()),
                    timeout=0.5,
                    return_when=concurrent.futures.FIRST_COMPLETED
                )

                for fut in done_futs:
                    fname = futures.pop(fut)
                    completed += 1
                    overall_bar.progress(
                        completed / len(files_to_run),
                        text=f"Completed {completed}/{len(files_to_run)} files…"
                    )
                    try:
                        rows, segs = fut.result()
                        all_rows_new.extend(rows)
                        new_results.extend(segs)
                        has_error = any(s.get("status") == "error" for s in segs)
                        file_status[fname] = "error" if has_error else "done"
                        # Track hash for dedup
                        for uf in files_to_run:
                            if uf.name == fname:
                                st.session_state.file_hashes[fname] = file_hash(uf.getvalue())
                    except Exception as e:
                        file_status[fname] = "error"
                        new_results.append({
                            "file": fname, "status": "error", "error": str(e)
                        })

                render_status()
                time.sleep(0.1)

        # Persist results
        st.session_state.all_rows.extend(all_rows_new)
        st.session_state.results.extend(new_results)

        overall_bar.progress(1.0, text="Batch complete!")
        time.sleep(0.5)
        overall_bar.empty()
        status_area.empty()
        st.rerun()

    # ── Per-file result cards ─────────────────────────────────────────────────
    if st.session_state.results:
        st.markdown("---")

        # Batch summary table
        by_file: dict[str, list] = defaultdict(list)
        for r in st.session_state.results:
            by_file[r["file"]].append(r)

        summary_rows = []
        for fname, segs in by_file.items():
            best_status = (
                "error" if all(s.get("status") == "error" for s in segs)
                else "needs_review" if any(s.get("status") in ("needs_review", "error") for s in segs)
                else "success"
            )
            confs = [s["confidence"] for s in segs if s.get("confidence")]
            summary_rows.append({
                "File": fname,
                "Segments": len(segs),
                "Status": best_status,
                "Avg Confidence": f"{sum(confs)/len(confs):.0%}" if confs else "—",
                "Rows": sum(s.get("rows", 0) for s in segs),
                "Multi-Invoice": "⚠️ Yes" if any(s.get("multi_invoice") for s in segs) else "No",
            })

        st.markdown('<div class="section-label">Batch Results</div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        # Expandable detail per file
        st.markdown("---")
        st.markdown('<div class="section-label">Per-File Details</div>', unsafe_allow_html=True)

        for fname, segs in by_file.items():
            multi = any(s.get("multi_invoice") for s in segs)
            n_err = sum(1 for s in segs if s.get("status") == "error")
            n_rev = sum(1 for s in segs if s.get("status") == "needs_review")

            label = f"📄 {fname}"
            if multi:
                label += f"  ·  ⚠️ {len(segs)} invoices"
            if n_err:
                label += f"  ·  ❌ {n_err} error(s)"
            if n_rev:
                label += f"  ·  ⚠️ {n_rev} review"

            with st.expander(label, expanded=(n_err > 0 or multi)):
                if multi:
                    st.markdown(
                        f'<div class="multi-banner">This PDF contains <strong>{len(segs)} separate invoice(s)</strong> — each extracted independently.</div>',
                        unsafe_allow_html=True
                    )

                for seg in segs:
                    conf       = seg.get("confidence", 0) or 0
                    bar_color  = "#4ade80" if conf >= threshold else "#fbbf24" if conf >= 0.6 else "#f87171"
                    pages_html = " ".join(
                        f'<span class="page-badge">p{p}</span>'
                        for p in seg.get("pages", [])
                    )
                    status_html = {
                        "success":      '<span class="pill-success">✓ success</span>',
                        "needs_review": '<span class="pill-warning">⚠ review</span>',
                        "error":        '<span class="pill-error">✗ error</span>',
                    }.get(seg.get("status", "error"), "")

                    ca, cb, cc, cd = st.columns([2, 3, 2, 2])
                    with ca:
                        st.markdown(f"**Split #{seg.get('split_num', 1)}** &nbsp; {status_html}", unsafe_allow_html=True)
                    with cb:
                        st.markdown(pages_html or "—", unsafe_allow_html=True)
                    with cc:
                        if conf:
                            st.markdown(
                                f'<div style="display:flex;align-items:center;gap:6px">'
                                f'<div class="conf-bar-wrap"><div class="conf-bar-fill" style="width:{int(conf*100)}%;background:{bar_color}"></div></div>'
                                f'<span style="font-family:\'DM Mono\',monospace;font-size:.78rem;color:{bar_color}">{conf:.0%}</span>'
                                f'</div>', unsafe_allow_html=True
                            )
                    with cd:
                        st.markdown(
                            f'<span style="font-family:\'DM Mono\',monospace;font-size:.78rem;color:#6b7898">'
                            f'inv# {seg.get("detected_inv", "—")}</span>',
                            unsafe_allow_html=True
                        )

                    for issue in seg.get("issues", []):
                        st.caption(f"⚠️ {issue}")
                    if seg.get("error"):
                        st.error(seg["error"])
                    if seg.get("extracted"):
                        st.json(seg["extracted"], expanded=False)

                    st.markdown(
                        "<hr style='border:none;border-top:1px solid #1e2330;margin:6px 0'>",
                        unsafe_allow_html=True
                    )


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — RESULTS
# ─────────────────────────────────────────────────────────────────────────────

with tab_results:
    if not st.session_state.all_rows:
        st.info("No data yet. Upload and extract invoices in the Extract tab.")
    else:
        csv_columns = build_csv_columns(config) if config else list(st.session_state.all_rows[0].keys())
        df = pd.DataFrame(st.session_state.all_rows)
        ordered = [c for c in csv_columns if c in df.columns]
        extra   = [c for c in df.columns if c not in ordered]
        df = df[ordered + extra]

        n_files  = df["File Name"].nunique() if "File Name" in df.columns else "—"
        n_splits = (
            df["Invoice Split #"].astype(str).ne("1").sum()
            if "Invoice Split #" in df.columns else 0
        )

        st.markdown(
            f"**{len(df)} rows** · **{n_files} file(s)**"
            + (f" · `{n_splits}` rows from split invoices" if n_splits else "")
        )

        # Filters
        f1, f2, f3, f4 = st.columns([2, 2, 1, 2])
        with f1:
            if "File Name" in df.columns:
                ff = st.multiselect("File", sorted(df["File Name"].unique()), placeholder="All files")
                if ff:
                    df = df[df["File Name"].isin(ff)]
        with f2:
            if "Invoice Split #" in df.columns:
                sf = st.multiselect("Split #", sorted(df["Invoice Split #"].dropna().unique()), placeholder="All")
                if sf:
                    df = df[df["Invoice Split #"].isin(sf)]
        with f3:
            if "Page(s)" in df.columns:
                pf = st.text_input("Page", placeholder="e.g. 2")
                if pf:
                    df = df[df["Page(s)"].astype(str).str.contains(pf.strip())]
        with f4:
            srch = st.text_input("Description search", placeholder="keyword…")
            if srch and "Description" in df.columns:
                df = df[df["Description"].str.contains(srch, case=False, na=False)]

        st.markdown("<br>", unsafe_allow_html=True)

        def highlight_split(row):
            if str(row.get("Invoice Split #", "1")) != "1":
                return ["background-color:#1c1400"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df.style.apply(highlight_split, axis=1),
            use_container_width=True, height=460, hide_index=True
        )
        st.caption("Amber rows = segments from split invoices (Invoice Split # > 1)")

        st.markdown("---")
        dl1, dl2, dl3 = st.columns(3)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        with dl1:
            st.download_button(
                "⬇  Download CSV",
                data=rows_to_csv_bytes(df.to_dict("records"), list(df.columns)),
                file_name=f"invoices_{ts}.csv", mime="text/csv",
                use_container_width=True
            )
        with dl2:
            buf = io.BytesIO()
            df.to_excel(buf, index=False, engine="openpyxl")
            st.download_button(
                "⬇  Download Excel", data=buf.getvalue(),
                file_name=f"invoices_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with dl3:
            st.download_button(
                "⬇  Download JSON",
                data=json.dumps(st.session_state.results, indent=2, default=str).encode(),
                file_name=f"invoices_{ts}.json", mime="application/json",
                use_container_width=True
            )

        # Needs review section
        review_segs = [r for r in st.session_state.results if r.get("status") == "needs_review"]
        if review_segs:
            st.markdown("---")
            st.markdown('<div class="section-label">Needs Human Review</div>', unsafe_allow_html=True)
            for r in review_segs:
                pages_str = ", ".join(f"p{p}" for p in r.get("pages", []))
                with st.expander(
                    f"⚠️ {r['file']} — split #{r.get('split_num', 1)} ({pages_str}) — {r.get('confidence', 0):.0%}"
                ):
                    for issue in r.get("issues", []):
                        st.caption(f"• {issue}")
                    if r.get("extracted"):
                        st.json(r["extracted"])


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — FIELD CONFIG
# ─────────────────────────────────────────────────────────────────────────────

with tab_config:
    st.markdown('<div class="section-label">fields_config.json — Live Editor</div>', unsafe_allow_html=True)
    st.caption("Edit fields, aliases, and CSV columns. Changes apply on next run.")

    if config:
        edited = st.text_area(
            "Config JSON", value=json.dumps(config, indent=2),
            height=500, label_visibility="collapsed"
        )
        ca, cb = st.columns(2)
        with ca:
            if st.button("✓ Validate JSON", use_container_width=True):
                try:
                    p = json.loads(edited)
                    miss = {"header_fields", "line_item_fields", "tax_line_items"} - set(p.keys())
                    if miss:
                        st.error(f"Missing sections: {miss}")
                    else:
                        st.success(
                            f"Valid — {len(p['header_fields'])} header fields, "
                            f"{len(p['line_item_fields'])} line-item fields"
                        )
                except json.JSONDecodeError as e:
                    st.error(f"JSON error: {e}")
        with cb:
            if st.button("💾 Save to File", use_container_width=True):
                try:
                    json.loads(edited)
                    with open(config_path, "w") as f:
                        f.write(edited)
                    st.success(f"Saved → {config_path}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Save failed: {e}")

        st.markdown("---")
        st.markdown("**CSV column order (auto-generated from config):**")
        st.code(" | ".join(build_csv_columns(config)), language=None)

        st.markdown("**Add a new header field:**")
        st.code('''{
  "po_number": {
    "csv_column": "PO Number",
    "aliases": ["po number", "purchase order", "po #"],
    "required": false
  }
}''', language="json")

        st.markdown("**Add a new tax label (extracted as separate line item):**")
        st.code('"known_tax_labels": ["gst", "cgst", "your_new_tax_here"]', language="json")
    else:
        st.error(f"Cannot load config from `{config_path}`")
