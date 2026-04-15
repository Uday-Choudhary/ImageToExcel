"""
ImageToExcel — Streamlit Web Application

Upload images of tables, invoices, and receipts to automatically extract
structured data and generate clean Excel files.

Powered by Llama 4 Vision (via Groq) + openpyxl.
Uses shared core modules for extraction and Excel generation.
"""

from __future__ import annotations

import logging
import time

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from core.constants import DEFAULT_VISION_MODEL
from core.excel_builder import build_excel_from_vision
from extractors.vision_extractor import VisionExtractor

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ── Load .env (works locally; on Streamlit Cloud use st.secrets) ──────────────
load_dotenv(override=True)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ImageToExcel — AI Table Extraction",
    page_icon=":material/bolt:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Premium Dark CSS ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

    /* ───── Global Reset ───── */
    html, body {
        font-family: 'Inter', sans-serif;
    }

    /* ───── App background ───── */
    .stApp {
        background: linear-gradient(135deg, #0a0f1e 0%, #0d1527 50%, #0a0f1e 100%) !important;
        background-attachment: fixed !important;
    }

    /* ───── Hide chrome ───── */
    footer { display: none !important; }

    /* ───── Main layout ───── */
    .block-container {
        padding: 3.5rem 2.5rem 3rem 2.5rem !important;
        max-width: 1280px !important;
    }

    /* ───── Headings ───── */
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Inter', sans-serif !important;
        color: #f1f5f9 !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
        line-height: 1.2 !important;
    }

    /* ───── Body text ───── */
    p, span, li,
    div[data-testid="stMarkdownContainer"] p,
    div[data-testid="stMarkdownContainer"] span {
        color: #94a3b8 !important;
        line-height: 1.7 !important;
    }

    small, .stCaption p, [data-testid="stCaptionContainer"] p {
        color: #64748b !important;
        font-size: 0.82rem !important;
    }

    /* ───── HERO SECTION ───── */
    .hero-wrap {
        background: linear-gradient(135deg,
            rgba(99,102,241,0.15) 0%,
            rgba(139,92,246,0.10) 40%,
            rgba(59,130,246,0.08) 100%);
        border: 1px solid rgba(99,102,241,0.25);
        border-radius: 20px;
        padding: 2.5rem 3rem;
        margin-top: 1rem;
        margin-bottom: 2.5rem;
        position: relative;
        overflow: hidden;
    }
    .hero-wrap::before {
        content: '';
        position: absolute;
        top: -60px;
        right: -60px;
        width: 260px;
        height: 260px;
        background: radial-gradient(circle, rgba(99,102,241,0.18) 0%, transparent 70%);
        pointer-events: none;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 7px;
        background: rgba(99,102,241,0.18);
        border: 1px solid rgba(99,102,241,0.35);
        color: #a5b4fc !important;
        font-size: 0.76rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        padding: 4px 12px;
        border-radius: 100px;
        margin-bottom: 1.2rem;
    }
    .hero-badge .dot {
        width: 6px;
        height: 6px;
        background: #6366f1;
        border-radius: 50%;
        animation: pulse-dot 2s infinite;
    }
    @keyframes pulse-dot {
        0%, 100% { opacity: 1; transform: scale(1); }
        50%       { opacity: 0.5; transform: scale(1.4); }
    }
    .hero-title {
        font-size: clamp(2rem, 5vw, 3rem) !important;
        font-weight: 800 !important;
        letter-spacing: -0.04em !important;
        background: linear-gradient(135deg, #f1f5f9 0%, #a5b4fc 60%, #818cf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin: 0 0 1rem 0 !important;
        line-height: 1.2 !important;
    }
    .hero-sub {
        font-size: 1.1rem !important;
        color: #94a3b8 !important;
        line-height: 1.7 !important;
        max-width: 600px;
        margin: 0 !important;
    }
    .hero-chips {
        display: flex;
        gap: 0.6rem;
        flex-wrap: wrap;
        margin-top: 1.8rem;
    }
    .chip {
        background: rgba(255,255,255,0.05);
        border: 1px solid rgba(255,255,255,0.1);
        color: #cbd5e1 !important;
        font-size: 0.8rem;
        font-weight: 500;
        padding: 4px 14px;
        border-radius: 100px;
    }

    /* ───── SIDEBAR ───── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #111827 0%, #0d1527 100%) !important;
        border-right: 1px solid rgba(99,102,241,0.15) !important;
    }
    [data-testid="stSidebar"] .block-container {
        padding: 1.8rem 1.5rem !important;
    }
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #f1f5f9 !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        margin-bottom: 0.25rem !important;
    }
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] li {
        color: #94a3b8 !important;
        font-size: 0.875rem !important;
        line-height: 1.6 !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,0.08) !important;
        margin: 1.25rem 0 !important;
    }
    .sidebar-logo {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 1.5rem;
    }
    .sidebar-logo-icon {
        width: 36px;
        height: 36px;
        background: linear-gradient(135deg, #6366f1, #8b5cf6);
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.1rem;
        flex-shrink: 0;
    }
    .sidebar-logo-text {
        font-size: 1.1rem !important;
        font-weight: 700 !important;
        color: #f1f5f9 !important;
        letter-spacing: -0.02em;
    }
    .step-item {
        display: flex;
        align-items: flex-start;
        gap: 12px;
        margin-bottom: 0.85rem;
    }
    .step-num {
        min-width: 24px;
        height: 24px;
        background: linear-gradient(135deg, #6366f1, #8b5cf6);
        border-radius: 50%;
        font-size: 0.72rem;
        font-weight: 700;
        color: #fff !important;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-top: 1px;
    }
    .step-text {
        font-size: 0.875rem !important;
        color: #94a3b8 !important;
        line-height: 1.5 !important;
    }
    .sidebar-model-badge {
        display: flex;
        align-items: center;
        gap: 8px;
        background: rgba(99,102,241,0.12);
        border: 1px solid rgba(99,102,241,0.25);
        border-radius: 8px;
        padding: 8px 12px;
        margin-top: 0.5rem;
    }
    .sidebar-model-badge .model-icon { font-size: 1rem; }
    .sidebar-model-badge .model-name {
        font-size: 0.7rem !important;
        font-weight: 600 !important;
        color: #a5b4fc !important;
        font-family: 'JetBrains Mono', monospace !important;
        word-break: break-all;
    }
    .format-tags {
        display: flex;
        gap: 6px;
        flex-wrap: wrap;
        margin-top: 0.4rem;
    }
    .format-tag {
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.1);
        color: #94a3b8 !important;
        font-size: 0.78rem;
        font-weight: 500;
        padding: 3px 10px;
        border-radius: 6px;
    }

    /* ───── SECTION HEADERS ───── */
    .section-header {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 1.25rem;
    }
    .section-header-icon {
        width: 32px;
        height: 32px;
        background: linear-gradient(135deg, rgba(99,102,241,0.25), rgba(139,92,246,0.15));
        border: 1px solid rgba(99,102,241,0.3);
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.9rem;
    }
    .section-header-text {
        font-size: 1.1rem !important;
        font-weight: 700 !important;
        color: #f1f5f9 !important;
        margin: 0 !important;
        line-height: 1 !important;
    }

    /* ───── GLASS CARD ───── */
    .glass-card {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 16px;
        padding: 1.75rem;
        backdrop-filter: blur(10px);
    }

    /* ───── METRIC CARDS ───── */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg,
            rgba(99,102,241,0.12) 0%,
            rgba(139,92,246,0.06) 100%) !important;
        border: 1px solid rgba(99,102,241,0.2) !important;
        border-radius: 14px !important;
        padding: 1.25rem 1rem !important;
        transition: border-color 0.2s ease, transform 0.2s ease;
        overflow-wrap: break-word !important;
    }
    [data-testid="stMetric"]:hover {
        border-color: rgba(99,102,241,0.45) !important;
        transform: translateY(-2px);
    }
    [data-testid="stMetricValue"] >div {
        color: #f1f5f9 !important;
        font-weight: 800 !important;
        font-size: clamp(1.4rem, 4vw, 1.8rem) !important;
        line-height: 1.2 !important;
        letter-spacing: -0.03em !important;
        white-space: normal !important;
        word-wrap: break-word !important;
    }
    [data-testid="stMetricLabel"] >div {
        color: #94a3b8 !important;
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
        margin-bottom: 0.3rem !important;
        white-space: normal !important;
    }
    [data-testid="stMetricDelta"] >div {
        font-size: 0.82rem !important;
    }

    /* ───── FILE UPLOADER ───── */
    [data-testid="stFileUploader"] {
        background: rgba(99,102,241,0.04) !important;
        border: 2px dashed rgba(99,102,241,0.3) !important;
        border-radius: 16px !important;
        padding: 1.5rem !important;
        transition: all 0.25s ease !important;
        box-shadow: 0 0 0 0 rgba(99,102,241,0) !important;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: rgba(99,102,241,0.65) !important;
        background: rgba(99,102,241,0.08) !important;
        box-shadow: 0 0 30px rgba(99,102,241,0.1) !important;
    }
    [data-testid="stFileUploader"] p {
        color: #64748b !important;
        font-size: 0.9rem !important;
    }
    [data-testid="stFileUploader"] small {
        color: #475569 !important;
    }

    /* ───── BUTTONS ───── */
    .stButton >button {
        background: rgba(255,255,255,0.05) !important;
        color: #e2e8f0 !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        border-radius: 10px !important;
        padding: 0.6rem 1.2rem !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        letter-spacing: 0.01em !important;
        transition: all 0.2s ease !important;
        backdrop-filter: blur(8px) !important;
    }
    .stButton >button:hover {
        background: rgba(255,255,255,0.09) !important;
        border-color: rgba(255,255,255,0.25) !important;
        transform: translateY(-1px) !important;
    }
    .stButton >button:active { transform: translateY(0) !important; }
    .stButton >button p { color: inherit !important; }

    /* Primary button */
    .stButton >button[kind="primary"] {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
        color: #ffffff !important;
        border: none !important;
        box-shadow: 0 4px 20px rgba(99,102,241,0.4) !important;
    }
    .stButton >button[kind="primary"]:hover {
        background: linear-gradient(135deg, #818cf8 0%, #a78bfa 100%) !important;
        box-shadow: 0 6px 28px rgba(99,102,241,0.55) !important;
        transform: translateY(-2px) !important;
    }
    .stButton >button[kind="primary"] p { color: #ffffff !important; }

    /* Download button */
    .stDownloadButton >button {
        background: linear-gradient(135deg, rgba(16,185,129,0.15), rgba(5,150,105,0.1)) !important;
        color: #6ee7b7 !important;
        border: 1px solid rgba(16,185,129,0.35) !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.65rem 1.2rem !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 2px 12px rgba(16,185,129,0.15) !important;
    }
    .stDownloadButton >button p { color: #6ee7b7 !important; }
    .stDownloadButton >button:hover {
        background: linear-gradient(135deg, rgba(16,185,129,0.25), rgba(5,150,105,0.18)) !important;
        border-color: rgba(16,185,129,0.6) !important;
        box-shadow: 0 4px 20px rgba(16,185,129,0.3) !important;
        transform: translateY(-1px) !important;
    }

    /* ───── TEXT INPUT ───── */
    .stTextInput >div >div >input {
        background: rgba(255,255,255,0.04) !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        border-radius: 8px !important;
        color: #e2e8f0 !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.875rem !important;
        padding: 0.6rem 0.9rem !important;
        transition: border-color 0.2s ease !important;
    }
    .stTextInput >div >div >input:focus {
        border-color: rgba(99,102,241,0.6) !important;
        box-shadow: 0 0 0 3px rgba(99,102,241,0.12) !important;
    }
    .stTextInput >div >div >input::placeholder {
        color: #475569 !important;
    }
    .stTextInput label {
        color: #94a3b8 !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
    }

    /* ───── ALERTS / STATUS ───── */
    div[class*="stAlert"] {
        border-radius: 10px !important;
        border: 1px solid !important;
        padding: 0.85rem 1.1rem !important;
        backdrop-filter: blur(8px) !important;
    }
    div[class*="stSuccess"] {
        background: rgba(16,185,129,0.1) !important;
        border-color: rgba(16,185,129,0.3) !important;
    }
    div[class*="stSuccess"] p { color: #6ee7b7 !important; }

    div[class*="stError"] {
        background: rgba(239,68,68,0.1) !important;
        border-color: rgba(239,68,68,0.3) !important;
    }
    div[class*="stError"] p { color: #fca5a5 !important; }

    div[class*="stWarning"] {
        background: rgba(251,191,36,0.08) !important;
        border-color: rgba(251,191,36,0.25) !important;
    }
    div[class*="stWarning"] p { color: #fcd34d !important; }

    div[class*="stInfo"] {
        background: rgba(99,102,241,0.1) !important;
        border-color: rgba(99,102,241,0.3) !important;
    }
    div[class*="stInfo"] p { color: #a5b4fc !important; }

    /* ───── PROGRESS BAR ───── */
    [data-testid="stProgress"] >div {
        background: rgba(255,255,255,0.07) !important;
        border-radius: 100px !important;
        height: 6px !important;
    }
    [data-testid="stProgress"] >div >div {
        background: linear-gradient(90deg, #6366f1, #8b5cf6) !important;
        border-radius: 100px !important;
        transition: width 0.4s ease !important;
    }
    .stProgress p {
        color: #94a3b8 !important;
        font-size: 0.82rem !important;
        margin-top: 0.4rem !important;
    }

    /* ───── TABS ───── */
    .stTabs [data-baseweb="tab-list"] {
        background: transparent !important;
        border-bottom: 1px solid rgba(255,255,255,0.08) !important;
        gap: 0 !important;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent !important;
        color: #64748b !important;
        font-weight: 600 !important;
        font-size: 0.875rem !important;
        padding: 0.75rem 1.25rem !important;
        border-radius: 0 !important;
        border-bottom: 2px solid transparent !important;
        transition: all 0.2s ease !important;
    }
    .stTabs [data-baseweb="tab"] p { color: inherit !important; }
    .stTabs [aria-selected="true"] {
        color: #818cf8 !important;
        border-bottom: 2px solid #6366f1 !important;
        background: transparent !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #a5b4fc !important;
        background: rgba(99,102,241,0.06) !important;
    }
    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 1.5rem !important;
    }

    /* ───── EXPANDER ───── */
    details {
        background: rgba(255,255,255,0.03) !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 10px !important;
        overflow: hidden !important;
        margin-top: 0.75rem !important;
    }
    details >summary {
        padding: 0.85rem 1.1rem !important;
        background: rgba(255,255,255,0.02) !important;
        cursor: pointer !important;
        transition: background 0.2s ease !important;
    }
    details >summary:hover { background: rgba(255,255,255,0.05) !important; }
    details >summary p {
        color: #94a3b8 !important;
        font-weight: 600 !important;
        font-size: 0.875rem !important;
    }
    details[open] >summary { border-bottom: 1px solid rgba(255,255,255,0.06) !important; }

    /* ───── DATAFRAME ───── */
    [data-testid="stDataFrame"] {
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 12px !important;
        overflow: hidden !important;
    }
    [data-testid="stDataFrame"] thead th {
        background: rgba(99,102,241,0.15) !important;
        color: #a5b4fc !important;
        font-size: 0.82rem !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
    }
    [data-testid="stDataFrame"] td {
        color: #cbd5e1 !important;
        font-size: 0.875rem !important;
    }

    /* ───── JSON VIEWER ───── */
    [data-testid="stJson"] {
        background: rgba(0,0,0,0.25) !important;
        border: 1px solid rgba(255,255,255,0.06) !important;
        border-radius: 10px !important;
        padding: 1rem !important;
    }

    /* ───── DIVIDER ───── */
    hr {
        border-color: rgba(255,255,255,0.07) !important;
        margin: 2rem 0 !important;
    }

    /* ───── IMAGE PREVIEW CARD ───── */
    .preview-card {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 12px;
        padding: 0.75rem;
        text-align: center;
        transition: border-color 0.2s ease, transform 0.2s ease;
    }
    .preview-card:hover {
        border-color: rgba(99,102,241,0.4);
        transform: translateY(-2px);
    }
    .preview-card img { border-radius: 8px; }
    .preview-filename {
        font-size: 0.78rem !important;
        color: #64748b !important;
        margin-top: 0.5rem !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }

    /* ───── EMPTY STATE ───── */
    .empty-state {
        text-align: center;
        padding: 3rem 2rem;
        background: rgba(255,255,255,0.02);
        border: 2px dashed rgba(255,255,255,0.07);
        border-radius: 20px;
    }
    .empty-state-icon {
        font-size: 3rem;
        line-height: 1;
        margin-bottom: 0.8rem;
    }
    .empty-state h3 {
        font-size: 1.35rem !important;
        font-weight: 700 !important;
        color: #475569 !important;
        margin-bottom: 0.6rem !important;
    }
    .empty-state p {
        font-size: 0.95rem !important;
        color: #334155 !important;
        max-width: 420px;
        margin: 0 auto !important;
        line-height: 1.6 !important;
    }

    /* ───── VALIDATION BADGE ───── */
    .val-badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 3px 10px;
        border-radius: 100px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    .val-pass { background: rgba(16,185,129,0.15); color: #6ee7b7; border: 1px solid rgba(16,185,129,0.3); }
    .val-fail { background: rgba(239,68,68,0.15); color: #fca5a5; border: 1px solid rgba(239,68,68,0.3); }

    /* ───── FOOTER ───── */
    .footer {
        text-align: center;
        padding: 1.5rem 0 0.5rem;
        border-top: 1px solid rgba(255,255,255,0.05);
        margin-top: 1rem;
    }
    .footer p {
        font-size: 0.82rem !important;
        color: #334155 !important;
        margin: 0 !important;
    }
    .footer a { color: #6366f1 !important; text-decoration: none !important; }
    .footer a:hover { color: #818cf8 !important; }

    /* ───── SPINNER ───── */
    .stSpinner >div { border-color: #6366f1 transparent transparent transparent !important; }

    /* ───── SCROLLBAR ───── */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: rgba(99,102,241,0.3); border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: rgba(99,102,241,0.5); }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <div class="sidebar-logo-icon">Ix</div>
        <span class="sidebar-logo-text">ImageToExcel</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### API Configuration")
    default_key = ""
    try:
        tmp_key = st.secrets.get("GROQ_API_KEY", "")
        if "your_key" not in tmp_key.lower():
            default_key = tmp_key
    except Exception:
        pass

    api_key_input = st.text_input(
        "Groq API Key",
        value=default_key,
        type="password",
        placeholder="gsk_...",
        help="Required if not set in Streamlit secrets or .env file.",
    )

    st.markdown("<hr>", unsafe_allow_html=True)

    st.markdown("**AI Engine**")
    st.markdown(f"""
    <div class="sidebar-model-badge">
        <span class="model-name">Llama 4 Vision</span>
    </div>
    <div style="margin-top:0.5rem; font-size:0.8rem; color:#64748b;">
        Powered by <a href="https://groq.com" target="_blank" style="color:#6366f1;text-decoration:none;font-weight:600;">Groq</a> ultra-fast inference
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("**How it works**")
    st.markdown("""
    <div class="step-item">
        <div class="step-num">1</div>
        <span class="step-text">Upload your image files (JPG, JPEG, PNG)</span>
    </div>
    <div class="step-item">
        <div class="step-num">2</div>
        <span class="step-text">Click <strong style="color:#a5b4fc;">Extract Data</strong>to analyze</span>
    </div>
    <div class="step-item">
        <div class="step-num">3</div>
        <span class="step-text">Review extracted tables &amp; download Excel</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("**Supported formats**")
    st.markdown("""
    <div class="format-tags">
        <span class="format-tag">JPG</span>
        <span class="format-tag">JPEG</span>
        <span class="format-tag">PNG</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="margin-top:1.5rem; padding: 10px 12px; background:rgba(99,102,241,0.08); border:1px solid rgba(99,102,241,0.18); border-radius:8px;">
        <p style="font-size:0.78rem !important; color:#64748b !important; margin:0 !important; line-height:1.5 !important;">
             Works with invoices, receipts, financial statements, data tables, and any structured image.
        </p>
    </div>
    """, unsafe_allow_html=True)


# ── Hero ───────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-wrap">
    <div class="hero-badge"><span class="dot"></span>AI-Powered Extraction</div>
    <h1 class="hero-title">Image → Excel</h1>
    <p class="hero-sub">
        Transform invoices, receipts, and data tables into clean, structured Excel files
        in seconds — powered by Llama 4 Vision via Groq.
    </p>
    <div class="hero-chips">
        <span class="chip">Ultra-fast inference</span>
        <span class="chip">Multi-sheet export</span>
        <span class="chip">Auto table detection</span>
        <span class="chip">✅ Math validation</span>
    </div>
</div>
""", unsafe_allow_html=True)


# ── Upload + Metrics Row ────────────────────────────────────────────────────────
col_upload, col_info = st.columns([3, 2], gap="large")

with col_upload:
    st.markdown("""
    <div class="section-header">
        <div class="section-header-icon"></div>
        <span class="section-header-text">Upload Documents</span>
    </div>
    """, unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Drag and drop files here",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

with col_info:
    st.markdown("""
    <div class="section-header">
        <div class="section-header-icon"></div>
        <span class="section-header-text">Overview</span>
    </div>
    """, unsafe_allow_html=True)
    m1, m2 = st.columns(2)
    m1.metric("Files", len(uploaded_files) if uploaded_files else 0)
    m2.metric("Status", "Ready " if uploaded_files else "Waiting")


# ── Image Preview ──────────────────────────────────────────────────────────────
if uploaded_files:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("""
    <div class="section-header">
        <div class="section-header-icon"></div>
        <span class="section-header-text">Preview</span>
    </div>
    """, unsafe_allow_html=True)

    preview_cols = st.columns(min(len(uploaded_files), 4))
    for i, f in enumerate(uploaded_files):
        with preview_cols[i % 4]:
            st.image(f, width="stretch")
            st.markdown(
                f'<p class="preview-filename" title="{f.name}">{f.name}</p>',
                unsafe_allow_html=True,
            )

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Extract Button ─────────────────────────────────────────────────────────
    col_btn, col_hint = st.columns([1, 3])
    with col_btn:
        run = st.button("Extract Data", type="primary", width="stretch")
    with col_hint:
        if not api_key_input:
            st.markdown("""
            <div style="padding:0.55rem 0.9rem; background:rgba(251,191,36,0.08);
                        border:1px solid rgba(251,191,36,0.2); border-radius:8px; margin-top:4px;">
                <p style="font-size:0.82rem !important; color:#fcd34d !important; margin:0 !important;">
                     No API key entered — will use key from <code style="background:rgba(255,255,255,0.08);
                    padding:1px 5px; border-radius:4px; color:#fcd34d;">.env</code>/ Streamlit secrets.
                </p>
            </div>
            """, unsafe_allow_html=True)

    if run:
        resolved_api_key = api_key_input
        if not resolved_api_key:
            try:
                resolved_api_key = st.secrets.get("GROQ_API_KEY")
            except Exception:
                pass
                
        extractor = VisionExtractor(api_key=resolved_api_key if resolved_api_key else None)

        results: list[tuple[str, dict]] = []
        all_raw: dict[str, dict] = {}

        progress_bar = st.progress(0, text="Initializing extraction…")
        status_area = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            pct = int((idx / len(uploaded_files)) * 100)
            progress_bar.progress(
                pct,
                text=f"Analyzing {uploaded_file.name}  ({idx + 1} of {len(uploaded_files)})"
            )

            with status_area.container():
                with st.spinner(f"Processing **{uploaded_file.name}** with Groq Vision…"):
                    image_bytes = uploaded_file.read()
                    t0 = time.time()
                    data = extractor.extract_from_image(
                        image_bytes=image_bytes,
                        filename=uploaded_file.name,
                    )
                    elapsed = time.time() - t0

            if data:
                sheet_name = uploaded_file.name.rsplit(".", 1)[0][:31]
                results.append((sheet_name, data))
                all_raw[uploaded_file.name] = data
                status_area.success(f"**{uploaded_file.name}** processed in {elapsed:.1f}s")
                logger.info("Extracted: %s (%.1fs)", uploaded_file.name, elapsed)
            else:
                status_area.error(f"Failed to extract data from **{uploaded_file.name}**")
                logger.warning("Failed: %s", uploaded_file.name)

        progress_bar.progress(100, text="Extraction complete")

        if results:
            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown("""
            <div class="section-header">
                <div class="section-header-icon"></div>
                <span class="section-header-text">Extracted Data</span>
            </div>
            """, unsafe_allow_html=True)

            tabs = st.tabs([name for name, _ in results])
            for tab, (sheet_name, data) in zip(tabs, results):
                with tab:
                    doc_sum  = data.get("document_summary", {})
                    entities = data.get("entities", {})
                    tables   = data.get("tables", [])

                    if doc_sum or entities:
                        c1, c2 = st.columns(2, gap="medium")
                        with c1:
                            if doc_sum:
                                st.markdown("""
                                <div class="section-header" style="margin-bottom:0.75rem;">
                                    <div class="section-header-icon" style="width:26px;height:26px;font-size:0.75rem;"></div>
                                    <span class="section-header-text" style="font-size:0.95rem;">Document Summary</span>
                                </div>
                                """, unsafe_allow_html=True)
                                st.json(doc_sum)
                        with c2:
                            if entities:
                                st.markdown("""
                                <div class="section-header" style="margin-bottom:0.75rem;">
                                    <div class="section-header-icon" style="width:26px;height:26px;font-size:0.75rem;"></div>
                                    <span class="section-header-text" style="font-size:0.95rem;">Entities</span>
                                </div>
                                """, unsafe_allow_html=True)
                                st.json(entities)

                    if tables:
                        for t_idx, t in enumerate(tables):
                            desc = t.get("table_description", f"Table {t_idx + 1}")
                            st.markdown(
                                f'<p style="font-size:0.95rem;font-weight:700;color:#e2e8f0;margin:1.25rem 0 0.6rem;">'
                                f' {desc}</p>',
                                unsafe_allow_html=True,
                            )
                            rows = t.get("rows", [])
                            if rows:
                                st.dataframe(
                                    pd.DataFrame(rows),
                                    width="stretch",
                                    hide_index=True,
                                )
                            else:
                                st.info("No row data detected in this table.")

                            val = t.get("validation", {})
                            if val:
                                check = val.get("math_check", "")
                                notes = val.get("notes", "")
                                passed = "pass" in check.lower()
                                badge_class = "val-pass" if passed else "val-fail"
                                icon = "" if passed else ""
                                st.markdown(
                                    f'<span class="val-badge {badge_class}">{icon} Validation: '
                                    f'{"Passed" if passed else "Failed"}</span>'
                                    f'<p style="font-size:0.8rem;color:#64748b;margin:2px 0 0;">{notes}</p>',
                                    unsafe_allow_html=True,
                                )
                    else:
                        st.warning(" No tabular data detected in this document.")

                    with st.expander("View Raw JSON Response", icon=":material/arrow_right:"):
                        st.json(data)

            # ── Excel Download ─────────────────────────────────────────────────
            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown("""
            <div class="section-header">
                <div class="section-header-icon"></div>
                <span class="section-header-text">Export</span>
            </div>
            """, unsafe_allow_html=True)

            with st.spinner("Generating Excel workbook…"):
                excel_bytes = build_excel_from_vision(results)

            st.success(f"Excel workbook ready — **{len(results)}** sheet(s) generated")

            dl_col, _ = st.columns([1, 3])
            with dl_col:
                st.download_button(
                    label="Download Excel File",
                    data=excel_bytes,
                    file_name="Extracted_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )

        else:
            st.error(
                "Data extraction failed for all uploaded files. "
                "Please verify your API key and try again."
            )

else:
    # ── Empty State ────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="empty-state">
        <div class="empty-state-icon"></div>
        <h3>No documents uploaded yet</h3>
        <p>Drag and drop your invoice, receipt, or table images into the upload
           zone above — or click to browse your files.</p>
    </div>
    """, unsafe_allow_html=True)


# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
    <p>
        ImageToExcel &nbsp;·&nbsp;
        Powered by <a href="https://groq.com" target="_blank">Groq</a>Llama 4 Vision &nbsp;·&nbsp;
        Built with <a href="https://streamlit.io" target="_blank">Streamlit</a>
    </p>
</div>
""", unsafe_allow_html=True)
