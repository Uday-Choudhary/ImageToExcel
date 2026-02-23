"""
ImageToExcel - Streamlit App
Upload images of tables/invoices and get a structured Excel file back.
Powered by Llama Vision (via Groq) + openpyxl.
"""

import os
import io
import json
import time
import base64
import tempfile

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from dotenv import load_dotenv
from groq import Groq

# ── Load .env (works locally; on Streamlit Cloud use st.secrets) ──────────────
load_dotenv()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ImageToExcel — AI Table Extractor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* Main gradient background */
    .stApp {
        background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
        min-height: 100vh;
    }

    /* Hero section */
    .hero {
        text-align: center;
        padding: 2.5rem 1rem 1.5rem;
    }
    .hero h1 {
        font-size: 3rem;
        font-weight: 700;
        background: linear-gradient(90deg, #a78bfa, #60a5fa, #34d399);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    .hero p {
        color: #94a3b8;
        font-size: 1.1rem;
        max-width: 600px;
        margin: 0 auto;
    }

    /* Glass cards */
    .glass-card {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 16px;
        padding: 1.5rem;
        backdrop-filter: blur(10px);
        margin-bottom: 1.2rem;
    }

    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: rgba(15, 12, 41, 0.85);
        border-right: 1px solid rgba(255,255,255,0.08);
    }
    [data-testid="stSidebar"] * { color: #e2e8f0 !important; }

    /* Upload zone */
    [data-testid="stFileUploader"] {
        background: rgba(167, 139, 250, 0.06) !important;
        border: 2px dashed rgba(167, 139, 250, 0.4) !important;
        border-radius: 12px !important;
        padding: 1rem !important;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: rgba(167, 139, 250, 0.8) !important;
    }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(90deg, #7c3aed, #2563eb) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.6rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        transition: all 0.2s ease !important;
        width: 100% !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(124,58,237,0.4) !important;
    }

    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(90deg, #059669, #0891b2) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        width: 100% !important;
        padding: 0.7rem 1.5rem !important;
        font-size: 1rem !important;
    }

    /* Status boxes */
    .stSuccess { border-radius: 10px !important; }
    .stError   { border-radius: 10px !important; }
    .stWarning { border-radius: 10px !important; }
    .stInfo    { border-radius: 10px !important; }

    /* Metric cards */
    [data-testid="stMetricValue"] { color: #a78bfa !important; font-weight: 700 !important; }
    [data-testid="stMetricLabel"] { color: #94a3b8 !important; }

    /* Expander */
    .streamlit-expanderHeader {
        background: rgba(255,255,255,0.04) !important;
        border-radius: 8px !important;
        color: #e2e8f0 !important;
    }

    /* Divider */
    hr { border-color: rgba(255,255,255,0.08) !important; }

    /* Tab styling */
    .stTabs [data-baseweb="tab"] { color: #94a3b8 !important; }
    .stTabs [aria-selected="true"] { color: #a78bfa !important; border-bottom-color: #a78bfa !important; }
</style>
""", unsafe_allow_html=True)

# ── Vision Pipeline (inline, no file I/O needed for Streamlit) ────────────────

VISION_PROMPT = """### INPUT ANALYSIS PHASE
1. **Identify Script Style**: First, detect if the text is printed, handwritten, or a mix. Adjust your internal recognition weights for handwritten characters.
2. **Spatial Anchoring**: Use visual lines and margins to define table boundaries.

### EXTRACTION REQUIREMENTS
1. **Hierarchical Extraction**:
   - **Header Data**: Extract titles, dates, and names.
   - **Tabular Structures**: Reconstruct grid-like data.
   - **Currency Detection**: Capture currency symbols exactly.
2. **Handwritten Edge Cases**:
   - Ignore crossed-out text.
   - Use [illegible] for unreadable words.
3. **The "Logical Audit"**:
   - Manually recalculate sums.
   - Report discrepancies as `handwritten_total` and `computed_total`.

### OUTPUT STRUCTURE (JSON)
Constraints:
- Return ONLY valid JSON.
- Do not include comments or markdown formatting (```json).
- "normalized_value" must be a single number or string, NOT a mathematical expression.

{
  "document_summary": { "style": "handwritten/printed", "domain": "auto-detect" },
  "entities": { "label": "value" },
  "tables": [
    {
      "table_description": "string",
      "headers": [],
      "rows": [
        { "column_name": "raw_text", "normalized_value": "numeric_or_string", "currency": "ISO_CODE" }
      ],
      "validation": {
        "math_check": "passed/failed",
        "notes": "string"
      }
    }
  ]
}"""


def get_groq_client():
    """Get Groq client using Streamlit secrets or environment variable."""
    # Try Streamlit secrets first (for Streamlit Cloud deployment)
    try:
        api_key = st.secrets["GROQ_API_KEY"]
    except (KeyError, FileNotFoundError):
        api_key = os.environ.get("GROQ_API_KEY", "")

    if not api_key:
        return None, "GROQ_API_KEY not found. Add it to Streamlit secrets or your .env file."
    return Groq(api_key=api_key), None


def encode_image_bytes(image_bytes: bytes) -> str:
    """Encode image bytes to base64 string."""
    return base64.b64encode(image_bytes).decode("utf-8")


def extract_data_from_image(image_bytes: bytes, filename: str) -> dict | None:
    """Send image to Llama Vision model and return structured JSON data."""
    client, err = get_groq_client()
    if err:
        st.error(f"🔑 API Error: {err}")
        return None

    # Detect MIME type
    ext = filename.lower().rsplit(".", 1)[-1]
    mime = "image/png" if ext == "png" else "image/jpeg"

    b64 = encode_image_bytes(image_bytes)
    image_data_url = f"data:{mime};base64,{b64}"

    try:
        completion = client.chat.completions.create(
            model="meta-llama/llama-4-maverick-17b-128e-instruct",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": VISION_PROMPT},
                        {"type": "image_url", "image_url": {"url": image_data_url}},
                    ],
                }
            ],
            temperature=0.1,
            response_format={"type": "json_object"},
            stream=False,
        )

        content = completion.choices[0].message.content
        return json.loads(content)

    except json.JSONDecodeError:
        st.error(f"❌ Model returned invalid JSON for **{filename}**.")
        return None
    except Exception as e:
        st.error(f"❌ API Error for **{filename}**: {e}")
        return None


# ── Excel Generation ──────────────────────────────────────────────────────────

def build_excel(results: list[dict]) -> bytes:
    """
    Given a list of (sheet_name, vision_data) tuples, build an Excel file
    and return its bytes.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, data in results:
            doc_summary = data.get("document_summary", {})
            entities = data.get("entities", {})
            tables = data.get("tables", [])

            # Build main table DataFrame
            main_table_df = pd.DataFrame()
            if tables:
                t1 = tables[0]
                rows = t1.get("rows", [])
                if rows:
                    main_table_df = pd.DataFrame(rows)
                else:
                    main_table_df = pd.DataFrame(columns=t1.get("headers", []))

            if main_table_df.empty:
                main_table_df = pd.DataFrame({"Message": ["No tables extracted"]})

            start_table_row = max(len(entities) + len(doc_summary) + 4, 6)
            safe_sheet = sheet_name[:31]  # Excel sheet name limit

            main_table_df.to_excel(writer, sheet_name=safe_sheet, index=False, startrow=start_table_row)
            ws = writer.sheets[safe_sheet]

            # Write metadata
            ws.cell(row=1, column=1, value="DOCUMENT SUMMARY").font = Font(bold=True, size=12)
            r = 2
            for k, v in doc_summary.items():
                ws.cell(row=r, column=1, value=k.title() + ":").font = Font(bold=True)
                ws.cell(row=r, column=2, value=str(v))
                r += 1

            r += 1
            ws.cell(row=r, column=1, value="ENTITIES").font = Font(bold=True, size=12)
            r += 1
            for k, v in entities.items():
                ws.cell(row=r, column=1, value=k.title() + ":").font = Font(bold=True)
                ws.cell(row=r, column=2, value=str(v))
                r += 1

            # Style table header
            header_row = start_table_row + 1
            for col_num, value in enumerate(main_table_df.columns.values):
                cell = ws.cell(row=header_row, column=col_num + 1)
                cell.value = value
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")

            # Additional tables
            if len(tables) > 1:
                current_row = header_row + len(main_table_df) + 3
                for t in tables[1:]:
                    ws.cell(row=current_row, column=1, value=t.get("table_description", "Table")).font = Font(bold=True, size=11)
                    current_row += 1
                    headers = t.get("headers", [])
                    t_rows = t.get("rows", [])
                    if not headers and t_rows:
                        headers = list(t_rows[0].keys())
                    for c_idx, h in enumerate(headers):
                        cell = ws.cell(row=current_row, column=c_idx + 1, value=h)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="6F819D", end_color="6F819D", fill_type="solid")
                    current_row += 1
                    for row_data in t_rows:
                        for c_idx, h in enumerate(headers):
                            ws.cell(row=current_row, column=c_idx + 1, value=str(row_data.get(h, "")))
                        current_row += 1
                    current_row += 2

            # Auto-size columns
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_len:
                            max_len = len(str(cell.value))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    return output.getvalue()


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## ⚙️ Settings")
    st.markdown("---")

    # API Key input (for users who want to paste it directly)
    api_key_input = st.text_input(
        "🔑 Groq API Key",
        type="password",
        placeholder="gsk_...",
        help="Required if not set in Streamlit secrets or .env file.",
    )
    if api_key_input:
        os.environ["GROQ_API_KEY"] = api_key_input

    st.markdown("---")
    st.markdown("### 📋 How It Works")
    steps = [
        "1️⃣ Upload one or more images",
        "2️⃣ Click **Extract & Generate Excel**",
        "3️⃣ Preview the extracted data",
        "4️⃣ Download your `.xlsx` file",
    ]
    for s in steps:
        st.markdown(s)

    st.markdown("---")
    st.markdown("### 🧠 Model")
    st.markdown("`meta-llama/llama-4-maverick-17b-128e-instruct`")
    st.markdown("Powered by **[Groq](https://groq.com)** inference.")

    st.markdown("---")
    st.markdown("### 📁 Supported Formats")
    st.markdown("`.jpg` · `.jpeg` · `.png`")


# ── Hero ───────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
    <h1>📊 ImageToExcel</h1>
    <p>Upload images of invoices, tables, or receipts and instantly get a clean, structured Excel file powered by AI vision.</p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ── Upload Section ─────────────────────────────────────────────────────────────

col_upload, col_info = st.columns([2, 1], gap="large")

with col_upload:
    st.markdown("### 📤 Upload Images")
    uploaded_files = st.file_uploader(
        "Drop your images here",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

with col_info:
    st.markdown("### 📈 Quick Stats")
    m1, m2 = st.columns(2)
    m1.metric("Images Uploaded", len(uploaded_files) if uploaded_files else 0)
    m2.metric("Ready to Extract", "✅" if uploaded_files else "—")

st.markdown("---")

# ── Image Preview ──────────────────────────────────────────────────────────────

if uploaded_files:
    st.markdown("### 🖼️ Preview")
    preview_cols = st.columns(min(len(uploaded_files), 4))
    for i, f in enumerate(uploaded_files):
        with preview_cols[i % 4]:
            st.image(f, caption=f.name, use_container_width=True)

    st.markdown("---")

    # ── Extract Button ─────────────────────────────────────────────────────────
    col_btn, _ = st.columns([1, 2])
    with col_btn:
        run = st.button("🚀 Extract & Generate Excel", use_container_width=True)

    if run:
        results = []
        all_raw = {}

        # Progress
        progress_bar = st.progress(0, text="Starting extraction...")
        status_area = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            pct = int((idx / len(uploaded_files)) * 100)
            progress_bar.progress(pct, text=f"Processing **{uploaded_file.name}** ({idx+1}/{len(uploaded_files)})...")

            with status_area.container():
                with st.spinner(f"🔍 Analysing `{uploaded_file.name}` with Llama Vision..."):
                    image_bytes = uploaded_file.read()
                    t0 = time.time()
                    data = extract_data_from_image(image_bytes, uploaded_file.name)
                    elapsed = time.time() - t0

            if data:
                sheet_name = uploaded_file.name.rsplit(".", 1)[0][:31]
                results.append((sheet_name, data))
                all_raw[uploaded_file.name] = data
                status_area.success(f"✅ `{uploaded_file.name}` extracted in {elapsed:.1f}s")
            else:
                status_area.error(f"❌ Failed to extract `{uploaded_file.name}`")

        progress_bar.progress(100, text="Done!")

        if results:
            st.markdown("---")
            st.markdown("### 📊 Extracted Data Preview")

            tabs = st.tabs([name for name, _ in results])
            for tab, (sheet_name, data) in zip(tabs, results):
                with tab:
                    # Document summary
                    doc_sum = data.get("document_summary", {})
                    entities = data.get("entities", {})
                    tables = data.get("tables", [])

                    c1, c2 = st.columns(2)
                    with c1:
                        if doc_sum:
                            st.markdown("**📄 Document Summary**")
                            st.json(doc_sum)
                    with c2:
                        if entities:
                            st.markdown("**🏷️ Entities**")
                            st.json(entities)

                    if tables:
                        for t_idx, t in enumerate(tables):
                            desc = t.get("table_description", f"Table {t_idx+1}")
                            st.markdown(f"**🗂️ {desc}**")
                            rows = t.get("rows", [])
                            if rows:
                                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                            else:
                                st.info("No rows found in this table.")

                            val = t.get("validation", {})
                            if val:
                                check = val.get("math_check", "")
                                notes = val.get("notes", "")
                                badge = "✅" if "pass" in check.lower() else "⚠️"
                                st.caption(f"{badge} Math check: **{check}** — {notes}")
                    else:
                        st.warning("No tables were detected in this image.")

                    # Raw JSON expander
                    with st.expander("🔍 View Raw JSON"):
                        st.json(data)

            # ── Excel Download ─────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### ⬇️ Download Excel")

            with st.spinner("Building Excel file..."):
                excel_bytes = build_excel(results)

            st.success(f"🎉 Excel file ready with **{len(results)}** sheet(s)!")

            dl_col, _ = st.columns([1, 2])
            with dl_col:
                st.download_button(
                    label="📥 Download Extracted_Data.xlsx",
                    data=excel_bytes,
                    file_name="Extracted_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        else:
            st.error("No data could be extracted from any of the uploaded images. Please check your API key and try again.")

else:
    # Empty state
    st.markdown("""
    <div style="text-align:center; padding: 3rem; color: #64748b;">
        <div style="font-size: 4rem; margin-bottom: 1rem;">📂</div>
        <h3 style="color: #94a3b8;">Upload images to get started</h3>
        <p>Supported formats: JPG, JPEG, PNG</p>
    </div>
    """, unsafe_allow_html=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#475569; font-size:0.85rem;'>"
    "ImageToExcel · Powered by Llama Vision on Groq · Built with Streamlit"
    "</p>",
    unsafe_allow_html=True,
)
