import os
import re
import io
import math
import time
import hashlib
import sqlite3
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st

try:
    from itsdangerous import URLSafeTimedSerializer
except Exception:
    URLSafeTimedSerializer = None

try:
    from streamlit_cookies_controller import CookieController
except Exception:
    CookieController = None
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ============================================================
# CONFIG
# ============================================================

APP_NAME = "Knowledge Base"
DATA_DIR = Path("knowledge_base_data")
PDF_DIR = DATA_DIR / "pdfs"
DB_PATH = DATA_DIR / "knowledge_base.db"

DATA_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)

st.set_page_config(
    page_title="Knowledge Base",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CSS
# ============================================================

st.markdown(
    """
<style>
:root {
    --navy: #0b2538;
    --navy-2: #123b50;
    --teal: #00a982;
    --teal-dark: #007f72;
    --green-soft: #e7f8f1;
    --bg: #f5f8fa;
    --white: #ffffff;
    --border: #d9e3e8;
    --text: #102d42;
    --muted: #687b87;
}

html, body, [class*="css"] {
    font-family: Arial, Helvetica, sans-serif;
}

.stApp {
    background: linear-gradient(180deg, #f8fbfc 0%, #f2f6f8 100%);
    color: var(--text);
}

[data-testid="stHeader"] { background: transparent; }

/* Hide Streamlit's default upper-right toolbar/menu icons.
   The app's own gear control remains visible because it is rendered in the page body. */
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"],
[data-testid="stMainMenu"],
button[kind="header"],
[data-testid="stHeaderActionElements"] {
    display: none !important;
    visibility: hidden !important;
}

/* Keep the header area clean after removing the native controls. */
[data-testid="stHeader"] {
    height: 0 !important;
    min-height: 0 !important;
}

/* Hide the default sidebar container; this app uses its own page navigation. */
[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"] {
    display: none !important;
}


.kb-topbar {
    display: flex;
    align-items: center;
    min-height: 76px;
    padding: 8px 10px 12px 8px;
    border-bottom: 1px solid #dfe7eb;
    background: linear-gradient(105deg, #ffffff 0%, #f7fbfc 70%, #e8f7f7 100%);
    margin-bottom: 8px;
}

.brand-block { width: 145px; }
.brand-mark {
    width: 46px;
    height: 7px;
    border: 3px solid #00a982;
    margin-bottom: 7px;
}
.brand-name { font-size: 14px; font-weight: 700; line-height: 1.05; color: #111; }
.brand-divider { height: 45px; width: 1px; background: #b8c7cf; margin: 0 22px 0 8px; }
.app-title { font-size: 28px; font-weight: 700; color: var(--text); line-height: 1; }
.app-subtitle { margin-top: 5px; font-size: 14px; color: #304a5c; }
.title-block { flex: 1; }
.top-tagline { text-align: right; color: #18394e; font-size: 12px; line-height: 1.3; margin-right: 12px; }
.top-actions { width: 36px; }
/* Header action group: Authorized User + admin gear stay together. */
.st-key-header_shell {
    position: relative !important;
    width: 100% !important;
    height: 76px !important;
    min-height: 76px !important;
    margin: 0 0 8px 0 !important;
    padding: 0 !important;
    overflow: visible !important;
}
.st-key-header_shell .kb-topbar {
    position: relative !important;
}
.st-key-gear_wrap {
    position: absolute !important;
    top: 50% !important;
    right: 10px !important;
    width: 30px !important;
    height: 30px !important;
    margin: 0 !important;
    padding: 0 !important;
    transform: translateY(-50%) !important;
    z-index: 10000 !important;
    pointer-events: none !important;
}
.st-key-gear_wrap > div,
.st-key-gear_wrap div[data-testid="stPopover"] {
    width: 30px !important;
    margin: 0 !important;
    padding: 0 !important;
    pointer-events: auto !important;
}
.st-key-gear_wrap div[data-testid="stPopover"] > button {
    width: 30px !important;
    min-width: 30px !important;
    height: 30px !important;
    min-height: 30px !important;
    padding: 0 !important;
    border: 1px solid #cfdde2 !important;
    border-radius: 6px !important;
    background: #e8f7f7 !important;
    box-shadow: none !important;
    color: #315468 !important;
    font-size: 13px !important;
    line-height: 30px !important;
}
.st-key-gear_wrap div[data-testid="stPopover"] > button:hover,
.st-key-gear_wrap div[data-testid="stPopover"] > button:focus {
    background: #e8f7f7 !important;
    color: #087c63 !important;
    box-shadow: none !important;
}


.bell { font-size: 22px; color: var(--navy); }

.exact-answer-card {
    background: linear-gradient(110deg, #f1fcf8, #ffffff 65%);
    border: 1px solid #7ad8bd;
    border-radius: 10px;
    padding: 18px 20px 16px;
    box-shadow: 0 3px 12px rgba(12, 54, 70, .05);
    margin-top: 8px;
}
.exact-answer-head { display: flex; justify-content: space-between; align-items: center; }
.exact-answer-title { color: #07866b; font-size: 19px; font-weight: 700; vertical-align: middle; }
.check-circle {
    display: inline-flex; width: 31px; height: 31px; border-radius: 50%;
    align-items: center; justify-content: center; background: #00a982; color: white;
    font-weight: 800; margin-right: 8px;
}
.match-pill { background: #d9f5ea; color: #087b64; font-weight: 700; padding: 5px 12px; border-radius: 20px; font-size: 12px; }
.exact-answer-note { margin: 5px 0 10px 39px; color: var(--muted); font-size: 12px; }
.exact-answer-text {
    margin: 0 0 12px 0; padding: 14px 18px; border-left: 4px solid var(--teal);
    background: rgba(255,255,255,.78); color: #172f42; font-size: 16px; line-height: 1.55;
}
.answer-meta { display: flex; flex-wrap: wrap; gap: 22px; color: #506672; font-size: 12px; padding-left: 2px; }

.source-header { background: white; border: 1px solid var(--border); border-bottom: 0; border-radius: 10px 10px 0 0; padding: 14px 16px; }
.source-title { font-size: 18px; font-weight: 700; color: var(--text); }
.source-meta { color: var(--muted); font-size: 12px; margin-top: 3px; }

.panel-title { font-size: 17px; font-weight: 700; color: var(--text); margin: 3px 0 8px; }
.related-title { color: #0561a0; font-weight: 700; font-size: 14px; }
.related-number { float: left; width: 23px; height: 23px; background: #dfe9ed; border-radius: 4px; text-align: center; line-height: 23px; font-weight: 700; color: #294a5c; }

.welcome-card {
    margin: 42px auto; max-width: 720px; text-align: center; background: white;
    border: 1px solid var(--border); border-radius: 14px; padding: 42px;
    box-shadow: 0 5px 18px rgba(12,54,70,.05);
}
.welcome-icon { font-size: 42px; color: var(--teal); }
.welcome-title { font-size: 25px; font-weight: 700; color: var(--text); margin-top: 8px; }
.welcome-text { color: var(--muted); max-width: 560px; margin: 10px auto; line-height: 1.6; font-size: 14px; }
.welcome-stats { display: flex; justify-content: center; gap: 35px; color: #57707e; margin-top: 18px; font-size: 12px; }

.page-heading { display:flex; justify-content:space-between; align-items:center; margin: 18px 0; }
.page-title { font-size: 26px; font-weight: 700; color: var(--text); }
.page-description { color: var(--muted); font-size: 13px; margin-top: 4px; }
.admin-badge { background:#e4f6f0; color:#087c63; font-size:11px; font-weight:700; border-radius:20px; padding:6px 12px; }
.admin-card { max-width:420px; margin:80px auto 20px; text-align:center; }
.admin-icon { font-size:40px; color:var(--teal); }
.admin-title { font-size:24px; font-weight:700; color:var(--text); }
.admin-subtitle { color:var(--muted); margin-top:5px; font-size:13px; }
.content-gap { height: 10px; }
.bottom-nav-spacer { height: 18px; }
.bottom-nav-label { text-align:center; color:#6c808b; font-size:10px; padding:4px 0 8px; }


.auth-shell { max-width: 620px; margin: 70px auto 22px; text-align: center; }
.auth-brand-mark { width: 55px; height: 8px; border: 3px solid #00a982; margin: 0 auto 10px; }
.auth-brand { font-size: 14px; font-weight: 700; color: #111; line-height: 1.05; }
.auth-title { margin-top: 26px; font-size: 31px; font-weight: 700; color: var(--text); }
.auth-subtitle { margin-top: 6px; color: var(--muted); font-size: 14px; }
.auth-card-title { font-size: 24px; font-weight: 700; color: var(--text); margin-top: 20px; }
.auth-switch { text-align:center; color:var(--muted); font-size:12px; margin:12px 0 6px; }
.top-user {
    color:#315468;
    font-size:12px;
    white-space:nowrap;
    background:#e8f7f7;
    padding:7px 10px;
    border-radius:4px;
    position:absolute;
    right:46px;
    top:50%;
    transform:translateY(-50%);
    z-index:2;
}

/* Compact search-result cards */
.st-key-search_result_best,
.st-key-search_result_1,
.st-key-search_result_2,
.st-key-search_result_3,
.st-key-search_result_4,
.st-key-search_result_5,
.st-key-search_result_6,
.st-key-search_result_7,
.st-key-search_result_8,
.st-key-search_result_9 {
    margin-bottom: 7px !important;
}

.st-key-search_result_best [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_1 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_2 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_3 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_4 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_5 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_6 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_7 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_8 [data-testid="stVerticalBlockBorderWrapper"],
.st-key-search_result_9 [data-testid="stVerticalBlockBorderWrapper"] {
    padding: 8px 10px !important;
    border-radius: 7px !important;
}

.result-label {
    color:#0561a0;
    font-weight:700;
    font-size:11px;
    line-height:1.2;
}
.result-filename {
    margin-top:3px;
    color:#0561a0;
    font-weight:700;
    font-size:12px;
    line-height:1.35;
    word-break:break-word;
}
.result-score {
    margin-top:3px;
    color:#687b87;
    font-size:10px;
}

.st-key-search_result_best button,
.st-key-search_result_1 button,
.st-key-search_result_2 button,
.st-key-search_result_3 button,
.st-key-search_result_4 button,
.st-key-search_result_5 button,
.st-key-search_result_6 button,
.st-key-search_result_7 button,
.st-key-search_result_8 button,
.st-key-search_result_9 button {
    min-height: 28px !important;
    height: 28px !important;
    padding: 2px 8px !important;
    font-size: 11px !important;
    margin-top: 4px !important;
}

.best-match-card { background:linear-gradient(110deg,#f1fcf8,#fff 70%); border:1px solid #7ad8bd; border-radius:10px; padding:16px 18px; margin-top:8px; box-shadow:0 3px 12px rgba(12,54,70,.05); }
.best-match-head { display:flex; justify-content:space-between; align-items:center; }
.pdf-badge { display:inline-flex; background:#e94b3c; color:white; font-weight:800; font-size:10px; border-radius:4px; padding:4px 6px; margin-right:7px; }
.best-match-title { color:#07866b; font-size:18px; font-weight:700; }
.best-match-file { color:#102d42; font-size:20px; font-weight:700; margin-top:8px; }
.best-match-meta { color:#687b87; font-size:12px; margin-top:4px; }
.source-page-label { font-size:12px; font-weight:700; color:#315468; margin:14px 0 6px; padding:6px 10px; background:#eef7f5; border-left:3px solid #00a982; border-radius:4px; }


/* Search / document reader redesign */
.search-results-heading {
    display:flex;
    justify-content:space-between;
    align-items:center;
    margin:18px 0 2px;
    color:#102d42;
    font-size:15px;
    font-weight:800;
}
.sort-chip {
    border:1px solid #d8e3e8;
    background:#fff;
    border-radius:7px;
    padding:7px 10px;
    font-size:10px;
    font-weight:600;
    color:#315468;
}

.search-count { color:#687b87; font-size:11px; margin:2px 0 8px; }
.reader-toolbar {
    background:#ffffff; border:1px solid var(--border); border-radius:9px;
    padding:10px 12px; margin-bottom:7px;
    box-shadow:0 1px 3px rgba(11,37,56,.03);
}
.reader-title { font-size:13px; font-weight:700; color:var(--text); line-height:1.3; word-break:break-word; }
.reader-meta { font-size:11px; color:var(--muted); margin-top:2px; }
.reader-match {
    background:#e7f8f1; border:1px solid #9bdcc8; color:#087c63;
    border-radius:5px; padding:5px 8px; font-size:10px; font-weight:700;
    display:inline-block; margin-top:5px;
}

.reader-page-indicator {
    text-align:right;
    color:#687b87;
    font-size:11px;
    line-height:30px;
    padding-right:4px;
}

/* PDF reader: arrow controls float over the page edges instead of taking a
   separate toolbar row. The shell is positioned so the arrows stay centered
   against the rendered PDF page. */
.st-key-pdf_viewer_shell {
    position:relative !important;
    overflow:visible !important;
    padding:0 !important;
}
.st-key-pdf_prev_wrap,
.st-key-pdf_next_wrap {
    position:absolute !important;
    top:50% !important;
    transform:translateY(-50%) !important;
    z-index:20 !important;
    width:42px !important;
}
.st-key-pdf_prev_wrap { left:10px !important; }
.st-key-pdf_next_wrap { right:10px !important; }
.st-key-pdf_prev_wrap button,
.st-key-pdf_next_wrap button {
    width:42px !important;
    min-width:42px !important;
    height:42px !important;
    min-height:42px !important;
    padding:0 !important;
    border-radius:50% !important;
    border:1px solid rgba(16,45,66,.18) !important;
    background:rgba(255,255,255,.94) !important;
    box-shadow:0 2px 10px rgba(12,54,70,.12) !important;
    color:#123b50 !important;
    font-size:28px !important;
    line-height:1 !important;
    transition:transform .12s ease, box-shadow .12s ease !important;
}
.st-key-pdf_prev_wrap button:hover:not(:disabled),
.st-key-pdf_next_wrap button:hover:not(:disabled) {
    transform:scale(1.06) !important;
    box-shadow:0 4px 14px rgba(12,54,70,.18) !important;
}
.st-key-pdf_prev_wrap button:disabled,
.st-key-pdf_next_wrap button:disabled {
    opacity:.35 !important;
}

/* Reduce excess vertical spacing around the rendered page so the reader
   feels faster and more continuous while scrolling. */
.st-key-pdf_viewer_shell [data-testid="stImage"] {
    margin-top:0 !important;
    margin-bottom:4px !important;
}
.ai-answer-card {
    background: linear-gradient(110deg, #f1fcf8, #ffffff 72%);
    border: 1px solid #9bdcc8;
    border-radius: 10px;
    padding: 14px 17px;
    margin: 8px 0 12px;
    box-shadow: 0 3px 12px rgba(12,54,70,.04);
}
.ai-answer-head {
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:12px;
}
.ai-answer-title {
    color:#087c63;
    font-size:16px;
    font-weight:700;
}
.ai-answer-badge {
    background:#d9f5ea;
    color:#087b64;
    font-size:10px;
    font-weight:700;
    border-radius:12px;
    padding:4px 8px;
    white-space:nowrap;
}
.ai-answer-body {
    margin-top:8px;
    color:#172f42;
    font-size:14px;
    line-height:1.55;
}
.ai-source-note {
    margin-top:8px;
    color:#687b87;
    font-size:10px;
}

.match-panel {
    background:#ffffff; border:1px solid var(--border); border-radius:8px;
    padding:10px; margin-top:10px;
}
.match-panel-title { font-size:13px; font-weight:700; color:var(--text); margin-bottom:7px; }
.match-item {
    background:#f7fafb; border:1px solid #e1e8ec; border-radius:6px;
    padding:7px 8px; margin-bottom:6px;
}
.match-item-page { color:#0561a0; font-size:10px; font-weight:700; }
.match-item-text { color:#385362; font-size:10px; line-height:1.35; margin-top:2px; }
.result-snippet {
    color:#536b78; font-size:10px; line-height:1.4; margin-top:5px;
    display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden;
}
.result-page { color:#687b87; font-size:10px; margin-top:3px; }
.result-score-pill {
    float:right; background:#e7f8f1; color:#087c63; border-radius:10px;
    padding:2px 6px; font-size:9px; font-weight:700;
}
.source-page-label { font-size:11px; font-weight:700; color:#315468; margin:8px 0 5px; padding:5px 8px; background:#eef7f5; border-left:3px solid #00a982; border-radius:4px; }

/* Interactive knowledge tiles — visual language follows the supplied mockup */
.explorer-title { font-size:15px; font-weight:800; color:#123b50; margin:14px 0 8px; }
.explorer-subtitle { font-size:10px; color:#687b87; margin-bottom:8px; }
.tile-grid { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; }
.knowledge-tile { background:#fff; border:1px solid #d9e3e8; border-radius:10px; padding:9px 12px; box-shadow:0 1px 2px rgba(11,37,56,.04); }
.knowledge-tile:hover { border-color:#00a982; box-shadow:0 2px 8px rgba(0,169,130,.10); }
.tile-name { font-size:11px; font-weight:700; color:#123b50; }
.tile-count { font-size:9px; color:#687b87; margin-top:2px; }

/* Four explorer summary cards */
.st-key-knowledge_mode_all button,
.st-key-knowledge_mode_category button,
.st-key-knowledge_mode_device button,
.st-key-knowledge_mode_topic button {
    min-height:66px !important;
    height:66px !important;
    padding:8px 12px !important;
    border-radius:9px !important;
    text-align:left !important;
    white-space:pre-line !important;
    font-size:13px !important;
    font-weight:700 !important;
    line-height:1.35 !important;
    border:1px solid #d8e3e8 !important;
    box-shadow:0 1px 3px rgba(11,37,56,.04) !important;
}
.st-key-knowledge_mode_all button[kind="primary"] {
    background:#18394d !important;
    border-color:#18394d !important;
    color:#fff !important;
}
.st-key-knowledge_mode_category button,
.st-key-knowledge_mode_device button,
.st-key-knowledge_mode_topic button {
    background:#fff !important;
    color:#17344a !important;
}
.st-key-knowledge_mode_category button:hover,
.st-key-knowledge_mode_device button:hover,
.st-key-knowledge_mode_topic button:hover {
    border-color:#00a982 !important;
    color:#087c63 !important;
}

.browse-heading {
    display:flex;
    justify-content:space-between;
    align-items:center;
    margin:16px 0 7px;
    color:#102d42;
    font-size:15px;
    font-weight:800;
}
.browse-viewall { color:#0067c5; font-size:11px; font-weight:700; }
.selected-explorer {
    margin:5px 0 8px;
    color:#087c63;
    background:#e8f8f3;
    border:1px solid #a8dfd0;
    border-radius:6px;
    padding:5px 8px;
    font-size:10px;
}

/* Six compact browse tiles per row */
.st-key-knowledge_tile_Category_0 button,
.st-key-knowledge_tile_Category_1 button,
.st-key-knowledge_tile_Category_2 button,
.st-key-knowledge_tile_Category_3 button,
.st-key-knowledge_tile_Category_4 button,
.st-key-knowledge_tile_Category_5 button,
.st-key-knowledge_tile_Device_0 button,
.st-key-knowledge_tile_Device_1 button,
.st-key-knowledge_tile_Device_2 button,
.st-key-knowledge_tile_Device_3 button,
.st-key-knowledge_tile_Device_4 button,
.st-key-knowledge_tile_Device_5 button,
.st-key-knowledge_tile_Topic_0 button,
.st-key-knowledge_tile_Topic_1 button,
.st-key-knowledge_tile_Topic_2 button,
.st-key-knowledge_tile_Topic_3 button,
.st-key-knowledge_tile_Topic_4 button,
.st-key-knowledge_tile_Topic_5 button {
    min-height:86px !important;
    height:86px !important;
    padding:7px 4px !important;
    border-radius:8px !important;
    white-space:pre-line !important;
    line-height:1.35 !important;
    font-size:10px !important;
    font-weight:700 !important;
}

/* Streamlit controls */
button[kind="primary"] { background: var(--teal) !important; border-color: var(--teal) !important; }
button[kind="primary"]:hover { background: var(--teal-dark) !important; }
[data-testid="stFileUploader"] { background: white; border-radius: 10px; border: 1px dashed #9ab1bc; }
[data-testid="stVerticalBlockBorderWrapper"] { border-color: var(--border) !important; border-radius: 9px !important; }
.stDownloadButton button { border-color: #00a982 !important; color: #087b64 !important; }

@media (max-width: 900px) {
    .brand-block { width: 110px; }
    .brand-divider, .top-tagline { display: none; }
    .app-title { font-size: 22px; }
    .app-subtitle { font-size: 12px; }
    .exact-answer-title { font-size: 16px; }
}
</style>
""",
    unsafe_allow_html=True,
)

# ============================================================
# DATABASE
# ============================================================

def db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA temp_store=MEMORY")
    conn.execute("PRAGMA cache_size=-16000")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            employee_id TEXT NOT NULL UNIQUE,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            stored_path TEXT NOT NULL,
            file_hash TEXT UNIQUE NOT NULL,
            category TEXT DEFAULT 'General',
            page_count INTEGER DEFAULT 0,
            file_size INTEGER DEFAULT 0,
            uploaded_at TEXT NOT NULL,
            indexed_at TEXT,
            status TEXT DEFAULT 'Indexed'
        );

        CREATE TABLE IF NOT EXISTS chunks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER NOT NULL,
            page_number INTEGER NOT NULL,
            chunk_index INTEGER NOT NULL,
            text TEXT NOT NULL,
            FOREIGN KEY(document_id) REFERENCES documents(id)
        );

        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_chunks_document_page ON chunks(document_id, page_number)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_chunks_document ON chunks(document_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_category ON documents(category)")
    conn.commit()
    conn.close()


init_db()

# ============================================================
# AUTHENTICATION
# ============================================================

def normalize_email(email):
    return email.strip().lower()


def hash_password(password, salt=None):
    if salt is None:
        salt = os.urandom(16)
    derived = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        210_000,
    )
    return salt.hex() + ":" + derived.hex()


def verify_password(password, stored):
    try:
        salt_hex, digest_hex = stored.split(":", 1)
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(digest_hex)
        actual = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            210_000,
        )
        return __import__("hmac").compare_digest(actual, expected)
    except Exception:
        return False


def create_user(first_name, last_name, employee_id, email, password):
    first_name = first_name.strip()
    last_name = last_name.strip()
    employee_id = employee_id.strip()
    email = normalize_email(email)

    if not all([first_name, last_name, employee_id, email, password]):
        return False, "All fields are required."

    if "@" not in email:
        return False, "Enter a valid email address."

    if len(password) < 8:
        return False, "Password must be at least 8 characters."

    conn = db()
    try:
        conn.execute(
            """
            INSERT INTO users
            (first_name, last_name, employee_id, email, password_hash, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                first_name,
                last_name,
                employee_id,
                email,
                hash_password(password),
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()
        return True, "Account created successfully. You can now sign in."
    except sqlite3.IntegrityError as e:
        message = str(e).lower()
        if "employee_id" in message:
            return False, "That Employee ID is already registered."
        if "email" in message:
            return False, "That email address is already registered."
        return False, "An account with those details already exists."
    finally:
        conn.close()


def authenticate_user(email, password):
    conn = db()
    row = conn.execute(
        "SELECT * FROM users WHERE email = ?",
        (normalize_email(email),),
    ).fetchone()
    conn.close()

    if not row or not verify_password(password, row["password_hash"]):
        return None

    return dict(row)


@st.cache_data(ttl=300, show_spinner=False)
def get_relevant_section(document_path, page_number, query, max_pages=4):
    """Extract the original PDF section related to the search query."""
    try:
        with fitz.open(document_path) as pdf:
            start_page = max(1, page_number)
            pages = []
            for pno in range(start_page, min(len(pdf), start_page + max_pages - 1) + 1):
                raw = pdf[pno - 1].get_text("text")
                if raw.strip():
                    pages.append((pno, raw))

            if not pages:
                return []

            terms = [t.lower() for t in re.findall(r"[A-Za-z0-9]+", query) if len(t) > 2]

            def is_heading(line):
                x = re.sub(r"\s+", " ", line).strip()
                if not x or len(x) > 120:
                    return False
                if re.match(r"^\d+(?:\.\d+)+\s", x):
                    return True
                if re.search(r"\b(steps?|checklist|procedure|process|requirements?|troubleshooting|instructions?|overview|guidelines?)\b", x, re.I):
                    return True
                letters = re.sub(r"[^A-Za-z]", "", x)
                return bool(letters) and letters.isupper() and len(letters) >= 4

            best_page = start_page
            best_line_index = 0
            best_score = -1
            for pno, raw in pages:
                lines = [re.sub(r"\s+", " ", x).strip() for x in raw.splitlines() if x.strip()]
                for i, line in enumerate(lines):
                    low = line.lower()
                    score = sum(low.count(t) for t in terms)
                    if score > best_score:
                        best_score = score
                        best_page = pno
                        best_line_index = i

            relevant = []
            found_heading = False
            for pno in range(best_page, min(len(pdf), best_page + max_pages - 1) + 1):
                raw = pdf[pno - 1].get_text("text")
                lines = [re.sub(r"\s+", " ", x).strip() for x in raw.splitlines() if x.strip()]
                if not lines:
                    continue

                if pno == best_page:
                    heading_idx = None
                    for i in range(min(best_line_index, len(lines) - 1), -1, -1):
                        if is_heading(lines[i]):
                            heading_idx = i
                            break
                    start_idx = heading_idx if heading_idx is not None else max(0, best_line_index)
                    found_heading = heading_idx is not None
                else:
                    start_idx = 0

                for i in range(start_idx, len(lines)):
                    line = lines[i]
                    if found_heading and i > start_idx and is_heading(line):
                        return _group_section_lines(relevant)
                    relevant.append((pno, line))

                if not found_heading:
                    break

            if len(relevant) < 2:
                raw = pdf[best_page - 1].get_text("text")
                return [(best_page, re.sub(r"\s+", " ", raw).strip())]
            return _group_section_lines(relevant)
    except Exception:
        return []


def _group_section_lines(relevant):
    grouped = []
    current_page = None
    current_lines = []
    for pno, line in relevant:
        if current_page is None:
            current_page = pno
        if pno != current_page:
            grouped.append((current_page, " ".join(current_lines)))
            current_page = pno
            current_lines = []
        current_lines.append(line)
    if current_page is not None and current_lines:
        grouped.append((current_page, " ".join(current_lines)))
    return grouped


@st.cache_data(ttl=60, show_spinner=False)
def get_document_by_id(document_id):
    conn = db()
    row = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
    conn.close()
    # Return plain Python data so Streamlit's cache can serialize the result.
    return dict(row) if row is not None else None


# ============================================================
# HELPERS
# ============================================================

def clean_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def make_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def split_text(text: str, chunk_size=1100, overlap=180):
    """
    Splits text approximately by words while preserving overlap.
    """
    words = text.split()
    if not words:
        return []

    chunks = []
    start = 0

    while start < len(words):
        end = min(len(words), start + chunk_size)
        chunk = " ".join(words[start:end]).strip()

        if chunk:
            chunks.append(chunk)

        if end >= len(words):
            break

        start = max(0, end - overlap)

    return chunks


def extract_pdf(pdf_bytes: bytes):
    pages = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page_number, page in enumerate(doc, start=1):
            text = clean_text(page.get_text("text"))
            pages.append((page_number, text))
        page_count = len(doc)
    return pages, page_count


def save_pdf(file_name: str, pdf_bytes: bytes, file_hash: str):
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", file_name)
    destination = PDF_DIR / f"{file_hash[:12]}_{safe_name}"
    destination.write_bytes(pdf_bytes)
    return str(destination)


def document_exists(file_hash):
    conn = db()
    row = conn.execute(
        "SELECT id FROM documents WHERE file_hash = ?",
        (file_hash,),
    ).fetchone()
    conn.close()
    return row


def add_document(file_name, pdf_bytes, category="General"):
    file_hash = make_hash(pdf_bytes)

    if document_exists(file_hash):
        return False, "This PDF has already been uploaded."

    try:
        pages, page_count = extract_pdf(pdf_bytes)
    except Exception as e:
        return False, f"Could not read PDF: {e}"

    stored_path = save_pdf(file_name, pdf_bytes, file_hash)

    conn = db()
    now = datetime.now().isoformat(timespec="seconds")

    cursor = conn.execute(
        """
        INSERT INTO documents
        (filename, stored_path, file_hash, category, page_count,
         file_size, uploaded_at, indexed_at, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            file_name,
            stored_path,
            file_hash,
            category,
            page_count,
            len(pdf_bytes),
            now,
            now,
            "Indexed",
        ),
    )

    document_id = cursor.lastrowid

    for page_number, page_text in pages:
        page_chunks = split_text(page_text)

        for chunk_index, chunk in enumerate(page_chunks):
            conn.execute(
                """
                INSERT INTO chunks
                (document_id, page_number, chunk_index, text)
                VALUES (?, ?, ?, ?)
                """,
                (
                    document_id,
                    page_number,
                    chunk_index,
                    chunk,
                ),
            )

    conn.commit()
    conn.close()

    return True, f"{file_name} indexed successfully."


@st.cache_data(ttl=30, show_spinner=False)
def get_documents():
    conn = db()
    rows = conn.execute(
        """
        SELECT *
        FROM documents
        ORDER BY uploaded_at DESC
        """
    ).fetchall()
    conn.close()
    # sqlite3.Row is not safely serializable by Streamlit's cache.
    return [dict(row) for row in rows]


@st.cache_data(ttl=60, show_spinner=False)
def get_knowledge_tiles():
    """Build interactive Category / Device / Topic facets from the locally indexed PDFs.

    Categories come directly from the document records. Devices and topics are
    detected from indexed filenames/text using a deterministic vocabulary; no AI
    service is required. Only terms actually present in the library are returned.
    """
    conn = db()
    docs = conn.execute("SELECT id, filename, category FROM documents ORDER BY filename").fetchall()
    chunks = conn.execute("SELECT document_id, text FROM chunks ORDER BY document_id, id").fetchall()
    conn.close()

    text_by_doc = {}
    for row in chunks:
        text_by_doc.setdefault(row["document_id"], []).append(row["text"] or "")

    # Practical HPE / enterprise support vocabulary. A tile is only shown when
    # the term actually occurs in a filename or indexed PDF text.
    device_terms = [
        "Aruba", "ClearPass", "ProLiant", "Alletra", "Nimble", "Primera",
        "3PAR", "MSA", "StoreEasy", "StoreOnce", "Synergy", "Apollo",
        "Superdome", "OneView", "iLO", "InfoSight", "HPE SimpliVity",
        "HPE GreenLake", "HPE Ezmeral", "HPE Networking", "HPE Switch",
        "HPE Router", "HPE Server", "HPE Storage", "HPE SAN", "HPE DL",
        "HPE ML", "HPE BL", "VMware", "FortiGate", "Cisco", "Switch",
        "Router", "Firewall", "Server", "Storage", "Printer", "Laptop",
        "Desktop", "Monitor", "Docking Station", "Access Point", "AP",
    ]

    topic_terms = [
        "Licensing", "Contracts", "Warranty", "Support", "Installation",
        "Configuration", "Troubleshooting", "Connectivity", "Networking",
        "Security", "Authentication", "Access", "Account", "Portal",
        "Firmware", "Software", "Hardware", "Drivers", "Updates",
        "Upgrade", "Migration", "Backup", "Recovery", "Storage",
        "Performance", "Incident", "Case", "Escalation", "Subscription",
        "Renewal", "Entitlement", "Serial Number", "Product Number",
        "Order", "Shipping", "Replacement", "RMA", "API", "SLA",
        "ClearPass", "Aruba", "OneView", "iLO", "VMware",
    ]

    def count_terms(terms):
        counts = {}
        for term in terms:
            needle = term.lower()
            count = 0
            for doc in docs:
                corpus = (doc["filename"] or "") + " " + " ".join(text_by_doc.get(doc["id"], []))
                if re.search(r"(?<![A-Za-z0-9])" + re.escape(needle) + r"(?![A-Za-z0-9])", corpus, re.I):
                    count += 1
            if count:
                counts[term] = count
        return sorted(counts.items(), key=lambda x: (-x[1], x[0].lower()))

    category_counts = {}
    for doc in docs:
        category = (doc["category"] or "General").strip() or "General"
        category_counts[category] = category_counts.get(category, 0) + 1

    return {
        "Category": sorted(category_counts.items(), key=lambda x: x[0].lower()),
        "Device": count_terms(device_terms),
        "Topic": count_terms(topic_terms),
    }


@st.cache_data(ttl=60, show_spinner=False)
def get_categories():
    conn = db()
    rows = conn.execute(
        """
        SELECT DISTINCT category
        FROM documents
        ORDER BY category
        """
    ).fetchall()
    conn.close()
    return [r["category"] for r in rows]


def get_all_chunks():
    conn = db()
    rows = conn.execute(
        """
        SELECT
            c.id,
            c.document_id,
            c.page_number,
            c.chunk_index,
            c.text,
            d.filename,
            d.category,
            d.stored_path
        FROM chunks c
        JOIN documents d ON d.id = c.document_id
        ORDER BY c.id
        """
    ).fetchall()
    conn.close()
    # Keep database rows as plain dictionaries for reliable caching/indexing.
    return [dict(row) for row in rows]


def delete_document(document_id):
    conn = db()

    row = conn.execute(
        "SELECT stored_path FROM documents WHERE id = ?",
        (document_id,),
    ).fetchone()

    if row:
        try:
            Path(row["stored_path"]).unlink(missing_ok=True)
        except Exception:
            pass

    conn.execute("DELETE FROM chunks WHERE document_id = ?", (document_id,))
    conn.execute("DELETE FROM documents WHERE id = ?", (document_id,))
    conn.commit()
    conn.close()


def format_bytes(value):
    if value is None:
        return "0 B"

    value = float(value)

    if value < 1024:
        return f"{value:.0f} B"
    if value < 1024**2:
        return f"{value / 1024:.1f} KB"
    if value < 1024**3:
        return f"{value / 1024**2:.1f} MB"

    return f"{value / 1024**3:.1f} GB"


def exact_passage(text, query, max_sentences=4, max_chars=1400):
    """Return verbatim text from the indexed PDF; never paraphrase."""
    normalized = re.sub(r"\s+", " ", text).strip()
    if not normalized:
        return ""

    sentences = re.split(r"(?<=[.!?])\s+", normalized)
    sentences = [s.strip() for s in sentences if s.strip()]

    terms = [
        t.lower()
        for t in re.findall(r"[A-Za-z0-9]+", query)
        if len(t) > 2
    ]

    if not terms or not sentences:
        return normalized[:max_chars]

    scored = []
    for i, sentence in enumerate(sentences):
        lower = sentence.lower()
        score = sum(lower.count(term) for term in terms)
        if score:
            scored.append((score, i))

    if not scored:
        return normalized[:max_chars]

    scored.sort(key=lambda x: (-x[0], x[1]))
    selected = set()

    for _, i in scored[:max_sentences]:
        selected.add(i)
        if len(selected) < max_sentences and i + 1 < len(sentences):
            selected.add(i + 1)

    passage = " ".join(sentences[i] for i in sorted(selected))

    if len(passage) > max_chars:
        passage = passage[:max_chars].rsplit(" ", 1)[0] + "..."

    return passage


def make_snippet(text, query, radius=260):
    text_clean = re.sub(r"\s+", " ", text).strip()
    if not query:
        return text_clean[:radius] + ("..." if len(text_clean) > radius else "")

    terms = [t.lower() for t in re.findall(r"\w+", query) if len(t) > 2]

    positions = []
    lower = text_clean.lower()

    for term in terms:
        pos = lower.find(term)
        if pos >= 0:
            positions.append(pos)

    if not positions:
        return text_clean[:radius] + ("..." if len(text_clean) > radius else "")

    center = min(positions)
    start = max(0, center - radius // 2)
    end = min(len(text_clean), start + radius)

    prefix = "..." if start > 0 else ""
    suffix = "..." if end < len(text_clean) else ""

    return prefix + text_clean[start:end] + suffix


# ============================================================
# SEARCH
# ============================================================

@st.cache_data(ttl=30, show_spinner=False)
def build_search_index():
    rows = get_all_chunks()

    if not rows:
        return None, [], []

    texts = [row["text"] for row in rows]

    vectorizer = TfidfVectorizer(
        lowercase=True,
        stop_words="english",
        ngram_range=(1, 2),
        min_df=1,
        max_df=0.98,
        sublinear_tf=True,
        dtype="float32",
    )

    matrix = vectorizer.fit_transform(texts)

    return vectorizer, matrix, [dict(r) for r in rows]


@st.cache_data(ttl=300, show_spinner=False)
def search_documents(query, category="All Categories", top_k=10):
    query = query.strip()

    # Category tiles can browse a category without requiring a keyword.
    if not query and category != "All Categories":
        conn = db()
        docs = conn.execute(
            "SELECT * FROM documents WHERE category = ? ORDER BY uploaded_at DESC LIMIT ?",
            (category, top_k),
        ).fetchall()
        results = []
        for doc in docs:
            chunk = conn.execute(
                "SELECT * FROM chunks WHERE document_id = ? ORDER BY page_number, chunk_index LIMIT 1",
                (doc["id"],),
            ).fetchone()
            if chunk:
                results.append({
                    "id": chunk["id"], "document_id": doc["id"],
                    "page_number": chunk["page_number"], "chunk_index": chunk["chunk_index"],
                    "text": chunk["text"], "filename": doc["filename"],
                    "category": doc["category"], "stored_path": doc["stored_path"],
                    "score": 0.0, "snippet": make_snippet(chunk["text"], ""),
                    "exact_passage": "",
                })
        conn.close()
        return results

    if not query:
        return []

    vectorizer, matrix, rows = build_search_index()

    if vectorizer is None:
        return []

    query_vector = vectorizer.transform([query])
    scores = cosine_similarity(query_vector, matrix).flatten()

    results = []

    for idx, score in enumerate(scores):
        row = rows[idx]

        if category != "All Categories" and row["category"] != category:
            continue

        if score <= 0:
            continue

        results.append(
            {
                **row,
                "score": float(score),
                "snippet": make_snippet(row["text"], query),
                "exact_passage": exact_passage(row["text"], query),
            }
        )

    # Each uploaded PDF can contain many indexed chunks/pages.
    # The search UI should show one result per PDF, not one result per chunk.
    # Keep the highest-scoring chunk from each PDF as that PDF's representative
    # result; the PDF viewer's "Matches in this PDF" section can still expose
    # the other matching pages from the selected document.
    best_by_document = {}
    for item in results:
        document_id = item["document_id"]
        current = best_by_document.get(document_id)
        if current is None or item["score"] > current["score"]:
            best_by_document[document_id] = item

    results = list(best_by_document.values())
    results.sort(key=lambda x: x["score"], reverse=True)

    return results[:top_k]


# PDF VIEWER
# ============================================================

@st.cache_data(ttl=600, show_spinner=False)
def render_pdf_page(document_path, page_number, scale=1.75):
    try:
        with fitz.open(document_path) as pdf:
            if page_number < 1 or page_number > len(pdf):
                return None
            page = pdf[page_number - 1]
            pix = page.get_pixmap(matrix=fitz.Matrix(float(scale), float(scale)), alpha=False)
            return pix.tobytes("png")
    except Exception:
        return None


@st.cache_data(ttl=120, show_spinner=False)
def find_document_matches(document_path, query, limit=8):
    """Find pages containing the user's actual search terms and return verbatim context."""
    try:
        terms = [t.lower() for t in re.findall(r"[A-Za-z0-9]+", query) if len(t) > 2]
        if not terms:
            return []
        matches = []
        with fitz.open(document_path) as pdf:
            for page_number, page in enumerate(pdf, start=1):
                raw = clean_text(page.get_text("text"))
                if not raw:
                    continue
                low = raw.lower()
                score = sum(low.count(term) for term in terms)
                if score <= 0:
                    continue
                matches.append({
                    "page": page_number,
                    "score": score,
                    "snippet": make_snippet(raw, query, radius=190),
                })
        matches.sort(key=lambda x: (-x["score"], x["page"]))
        return matches[:limit]
    except Exception:
        return []


# ============================================================
# UI STATE
# ============================================================

if "page" not in st.session_state:
    st.session_state.page = "Search"

if "selected_document" not in st.session_state:
    st.session_state.selected_document = None

if "selected_page" not in st.session_state:
    st.session_state.selected_page = 1

if "search_query" not in st.session_state:
    st.session_state.search_query = ""

if "knowledge_tile_type" not in st.session_state:
    st.session_state.knowledge_tile_type = "Category"

if "knowledge_tile_value" not in st.session_state:
    st.session_state.knowledge_tile_value = None

if "admin_authenticated" not in st.session_state:
    st.session_state.admin_authenticated = False



# ============================================================
# PDF HIGHLIGHTING
# ============================================================

@st.cache_data(ttl=600, show_spinner=False)
def render_pdf_page_highlighted(document_path, page_number, query="", scale=1.75):
    """Render a PDF page with matching query terms highlighted at the requested scale."""
    try:
        with fitz.open(document_path) as pdf:
            if page_number < 1 or page_number > len(pdf):
                return None
            page = pdf[page_number - 1]
            terms = [t for t in re.findall(r"[A-Za-z0-9]+", query) if len(t) > 2]
            highlighted = set()
            for term in terms[:12]:
                try:
                    for rect in page.search_for(term):
                        key = (round(rect.x0, 1), round(rect.y0, 1), round(rect.x1, 1), round(rect.y1, 1))
                        if key in highlighted:
                            continue
                        highlighted.add(key)
                        annot = page.add_highlight_annot(rect)
                        annot.update()
                except Exception:
                    continue
            pix = page.get_pixmap(matrix=fitz.Matrix(float(scale), float(scale)), alpha=False)
            return pix.tobytes("png")
    except Exception:
        return None


def clear_knowledge_caches():
    """Invalidate cached database/search/PDF-derived data after document changes."""
    for fn in (build_search_index, search_documents, get_documents, get_categories, get_document_by_id, get_relevant_section, render_pdf_page, render_pdf_page_highlighted):
        try:
            fn.clear()
        except Exception:
            pass


# ============================================================
# ADMIN AUTHENTICATION
# ============================================================

def get_admin_pin():
    """Read the admin PIN from Streamlit secrets first, then environment."""
    try:
        pin = st.secrets.get("ADMIN_PIN")
        if pin:
            return str(pin)
    except Exception:
        pass

    return os.getenv("ADMIN_PIN", "")


def admin_is_configured():
    return bool(get_admin_pin())


# ============================================================
# AUTH STATE
# ============================================================

if "access_authorized" not in st.session_state:
    st.session_state.access_authorized = False

if "cookie_restore_checked" not in st.session_state:
    st.session_state.cookie_restore_checked = False

if "page" not in st.session_state:
    st.session_state.page = "Search"

if "search_query" not in st.session_state:
    st.session_state.search_query = ""

if "selected_document" not in st.session_state:
    st.session_state.selected_document = None

if "selected_page" not in st.session_state:
    st.session_state.selected_page = 1

if "selected_result_id" not in st.session_state:
    st.session_state.selected_result_id = None

if "search_results" not in st.session_state:
    st.session_state.search_results = []

if "search_signature" not in st.session_state:
    st.session_state.search_signature = None

if "admin_authenticated" not in st.session_state:
    st.session_state.admin_authenticated = False


# ============================================================
# ============================================================
# PERSISTENT BROWSER ACCESS
# ============================================================
# Streamlit Cloud can lose client-side cookies across a fresh WebSocket
# connection. To make F5/refresh deterministic, the signed authorization
# token is persisted in the app URL. The token contains no user information.
# IMPORTANT: anyone who has the full authorized URL can access the app.

ACCESS_CODE = str(
    st.secrets.get("ACCESS_CODE", os.getenv("ACCESS_CODE", ""))
).strip()

TOKEN_SECRET = str(
    st.secrets.get("TOKEN_SECRET", os.getenv("TOKEN_SECRET", ""))
).strip()

if not TOKEN_SECRET:
    TOKEN_SECRET = hashlib.sha256(
        f"{os.getcwd()}::{os.getenv('HOSTNAME', 'streamlit')}".encode()
    ).hexdigest()


def get_token_serializer():
    if URLSafeTimedSerializer is None or not TOKEN_SECRET:
        return None
    return URLSafeTimedSerializer(
        TOKEN_SECRET,
        salt="knowledge-base-browser-access",
    )


def create_browser_token():
    serializer = get_token_serializer()
    if serializer is None:
        return ""
    return serializer.dumps({"authorized": True})


def validate_browser_token(token):
    if not token:
        return False
    serializer = get_token_serializer()
    if serializer is None:
        return False
    try:
        payload = serializer.loads(str(token))
        return bool(payload.get("authorized"))
    except Exception:
        return False


def get_url_access_token():
    try:
        return st.query_params.get("kb_access", "")
    except Exception:
        return ""


def browser_is_authorized():
    if st.session_state.access_authorized:
        return True

    token = get_url_access_token()
    if validate_browser_token(token):
        st.session_state.access_authorized = True
        return True

    return False


def authorize_browser():
    token = create_browser_token()
    if not token:
        return False

    # Query parameters survive a normal browser refresh on Streamlit Cloud.
    st.query_params["kb_access"] = token
    st.session_state.access_authorized = True
    return True


def clear_browser_access():
    st.session_state.access_authorized = False
    try:
        st.query_params.clear()
    except Exception:
        pass


def render_access_gate():
    st.markdown(
        """
        <div class="auth-shell">
            <div class="auth-brand-mark"></div>
            <div class="auth-brand">Hewlett Packard Enterprise</div>
            <div class="auth-title">Knowledge Base</div>
            <div class="auth-subtitle">
                Secure access to your organization's PDF knowledge base.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not ACCESS_CODE:
        st.error(
            "Access control is not configured. Add ACCESS_CODE to Streamlit Secrets."
        )
        st.stop()

    _, center, _ = st.columns([1, 2, 1])

    with center:
        st.markdown(
            '<div class="auth-card-title">Enter Access Code</div>',
            unsafe_allow_html=True,
        )
        st.caption("You only need to enter the code once on this browser.")

        with st.form("access_code_form"):
            entered_code = st.text_input(
                "Access Code",
                type="password",
                placeholder="Enter access code",
            )
            submitted = st.form_submit_button(
                "Access Knowledge Base",
                type="primary",
                use_container_width=True,
            )

        if submitted:
            if entered_code.strip() == ACCESS_CODE:
                if authorize_browser():
                    # No Continue button. The signed token is placed in the
                    # URL and the app immediately loads the Knowledge Base.
                    st.rerun()
                else:
                    st.error("Unable to create the browser authorization token.")
            else:
                st.error("Invalid access code.")

        st.caption(
            "The access code is never stored in the URL. A signed authorization "
            "token is used to keep this browser authorized."
        )


if not browser_is_authorized():
    render_access_gate()
    st.stop()


# HEADER
# ============================================================

user_name = "Authorized User"

with st.container(key="header_shell"):
    st.markdown(
        f"""
        <div class="kb-topbar">
            <div class="brand-block">
                <div class="brand-mark"></div>
                <div class="brand-name">Hewlett Packard<br>Enterprise</div>
            </div>
            <div class="brand-divider"></div>
            <div class="title-block">
                <div class="app-title">Knowledge Base</div>
                <div class="app-subtitle">Search. Find. Resolve.</div>
            </div>
            <div class="top-user">{user_name}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Gear is inside the same header container as Authorized User, so its vertical
    # position is tied to the header rather than the browser viewport.
    with st.container(key="gear_wrap"):
        with st.popover("⚙", use_container_width=False):
            st.markdown("**Knowledge Base Access**")
            st.caption(
                "Browser authorization: Persistent until manually cleared"
            )
            st.divider()

            if st.session_state.admin_authenticated:
                if st.button("Manage Documents", use_container_width=True):
                    st.session_state.page = "Manage Documents"
                    st.rerun()

                if st.button("Sign out admin", use_container_width=True):
                    st.session_state.admin_authenticated = False
                    st.session_state.page = "Search"
                    st.rerun()
            else:
                if st.button("🔒 Manage Documents", use_container_width=True):
                    st.session_state.page = "Admin Login"
                    st.rerun()

            if st.button("Clear Browser Access", use_container_width=True):
                clear_browser_access()
                st.session_state.admin_authenticated = False
                st.session_state.page = "Search"
                st.rerun()


# ============================================================
# ADMIN LOGIN
# ============================================================

if st.session_state.page == "Admin Login":
    st.markdown(
        """
        <div class="admin-card">
            <div class="admin-icon">⚙</div>
            <div class="admin-title">Admin Access</div>
            <div class="admin-subtitle">Enter the administrator PIN to manage knowledge-base documents.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not admin_is_configured():
        st.error("Admin access is not configured. Set ADMIN_PIN in Streamlit Secrets or an environment variable.")
    else:
        with st.form("admin_login_form"):
            pin = st.text_input("Admin PIN", type="password", placeholder="Enter admin PIN")
            submitted = st.form_submit_button("Unlock", type="primary", use_container_width=True)
        if submitted:
            if pin == get_admin_pin():
                st.session_state.admin_authenticated = True
                st.session_state.page = "Manage Documents"
                st.rerun()
            else:
                st.error("Incorrect admin PIN.")

    if st.button("← Back to Search", use_container_width=True):
        st.session_state.page = "Search"
        st.rerun()


# ============================================================
# ADMIN: MANAGE DOCUMENTS
# ============================================================

elif st.session_state.page == "Manage Documents":
    if not st.session_state.admin_authenticated:
        st.session_state.page = "Admin Login"
        st.rerun()

    st.markdown(
        """
        <div class="page-heading">
            <div>
                <div class="page-title">Manage Documents</div>
                <div class="page-description">Upload, manage, and index PDF documents for the knowledge base.</div>
            </div>
            <div class="admin-badge">ADMIN ONLY</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    upload_col, library_col = st.columns([0.95, 1.35], gap="large")

    with upload_col:
        st.markdown('<div class="panel-title">Upload Documents</div>', unsafe_allow_html=True)
        category = st.selectbox(
            "Category",
            ["General", "Policies", "Procedures", "Technical Support", "Licensing", "Training", "Product", "Account Management", "Other"],
            key="admin_category",
        )
        uploaded_files = st.file_uploader(
            "Drag and drop PDF files here",
            type=["pdf"],
            accept_multiple_files=True,
            key="admin_uploader",
        )
        if uploaded_files:
            st.caption(f"{len(uploaded_files)} PDF file(s) selected")
            for f in uploaded_files:
                st.write(f"📄 {f.name} · {format_bytes(len(f.getvalue()))}")

        if st.button("Upload & Index Documents", type="primary", use_container_width=True):
            if not uploaded_files:
                st.warning("Select at least one PDF file.")
            else:
                progress = st.progress(0)
                success_count = 0
                for i, uploaded_file in enumerate(uploaded_files):
                    ok, message = add_document(uploaded_file.name, uploaded_file.getvalue(), category)
                    if ok:
                        success_count += 1
                        st.success(message)
                    else:
                        st.warning(message)
                    progress.progress((i + 1) / len(uploaded_files))
                clear_knowledge_caches()
                st.session_state.search_results = []
                st.session_state.search_signature = None
                st.success(f"Completed. {success_count} document(s) indexed.")

    with library_col:
        st.markdown('<div class="panel-title">Document Library</div>', unsafe_allow_html=True)
        docs = get_documents()
        lc1, lc2 = st.columns([1.5, 1])
        with lc1:
            library_search = st.text_input("Search documents", placeholder="Search documents...", label_visibility="collapsed", key="library_search")
        with lc2:
            library_category = st.selectbox("Library category", ["All Categories"] + get_categories(), label_visibility="collapsed", key="library_category")

        filtered_docs = []
        for doc in docs:
            if library_search and library_search.lower() not in doc["filename"].lower():
                continue
            if library_category != "All Categories" and doc["category"] != library_category:
                continue
            filtered_docs.append(doc)

        if not filtered_docs:
            st.info("No documents match the current filters.")
        else:
            for doc in filtered_docs:
                with st.container(border=True):
                    a, b = st.columns([4, 1])
                    with a:
                        st.markdown(f"**📄 {doc['filename']}**")
                        st.caption(f"{doc['category']} · {doc['page_count']} pages · {format_bytes(doc['file_size'])} · ✓ {doc['status']}")
                    with b:
                        if st.button("Delete", key=f"admin_delete_{doc['id']}"):
                            delete_document(doc["id"])
                            clear_knowledge_caches()
                            st.session_state.search_results = []
                            st.session_state.search_signature = None
                            st.rerun()

    if st.button("← Back to Search"):
        st.session_state.page = "Search"
        st.rerun()


# ============================================================
# SEARCH KNOWLEDGE BASE
# ============================================================

else:
    st.session_state.page = "Search"

    # --------------------------------------------------------
    # SEARCH BAR — kept at the very top, matching the reference
    # --------------------------------------------------------
    with st.form("main_search_form", clear_on_submit=False):
        search_col, button_col = st.columns([7.2, 1.0], gap="small")
        with search_col:
            query = st.text_input(
                "Search",
                value=st.session_state.search_query,
                placeholder="Search the knowledge base...",
                label_visibility="collapsed",
                key="main_search_box",
            )
        with button_col:
            search_clicked = st.form_submit_button(
                "⌕  Search",
                type="primary",
                use_container_width=True,
            )

    if search_clicked:
        st.session_state.search_query = query
        st.session_state.selected_result_id = None
        st.session_state.search_signature = None

    active_query = st.session_state.search_query.strip()
    category = st.session_state.get("search_category", "All Categories")
    top_k = st.session_state.get("search_top_k", 10)

    # --------------------------------------------------------
    # TOP EXPLORER SWITCHER — All / Category / Device / Topic
    # --------------------------------------------------------
    tile_data = get_knowledge_tiles()
    docs = get_documents()
    total_docs = len(docs)
    tile_type = st.session_state.get("knowledge_tile_type", "Category")

    mode_cols = st.columns(4, gap="small")
    mode_defs = [
        ("All", total_docs),
        ("Category", len(tile_data.get("Category", []))),
        ("Device", len(tile_data.get("Device", []))),
        ("Topic", len(tile_data.get("Topic", []))),
    ]

    for idx, (label, count) in enumerate(mode_defs):
        with mode_cols[idx]:
            is_active = label == "All"
            if label == "All":
                mode_text = f"▱  All\n{count} PDFs"
            elif label == "Category":
                mode_text = f"▰  Category\n{count} categories"
            elif label == "Device":
                mode_text = f"▣  Device\n{count} devices"
            else:
                mode_text = f"▤  Topic\n{count} topics"

            if st.button(
                mode_text,
                key=f"knowledge_mode_{label.lower()}",
                use_container_width=True,
                type="primary" if is_active else "secondary",
            ):
                if label == "All":
                    st.session_state.knowledge_tile_type = "Category"
                    st.session_state.knowledge_tile_value = None
                    st.session_state.search_category = "All Categories"
                    st.session_state.search_query = ""
                else:
                    st.session_state.knowledge_tile_type = label
                    st.session_state.knowledge_tile_value = None
                st.session_state.knowledge_tile_show_all = False
                st.session_state.selected_result_id = None
                st.session_state.search_signature = None
                st.rerun()

    # Keep filtering state available to the search engine without adding a
    # visible sidebar or filter row to the reference layout.
    if "search_category" not in st.session_state:
        st.session_state.search_category = "All Categories"
    if "search_top_k" not in st.session_state:
        st.session_state.search_top_k = 10

    category = st.session_state.get("search_category", "All Categories")
    top_k = st.session_state.get("search_top_k", 10)

    # A selected Category tile is a category browse. Device/Topic tiles are
    # keyword shortcuts that use the normal TF-IDF PDF search.
    selected_tile = st.session_state.get("knowledge_tile_value")
    if selected_tile:
        st.markdown(
            f'<div class="selected-explorer">Selected {tile_type.lower()}: <b>{selected_tile}</b></div>',
            unsafe_allow_html=True,
        )

    should_search = bool(active_query) or (tile_type == "Category" and category != "All Categories")

    if should_search:
        search_signature = (active_query, category, top_k)
        if st.session_state.get("search_signature") != search_signature:
            st.session_state.search_results = search_documents(
                active_query,
                category=category,
                top_k=top_k,
            )
            st.session_state.search_signature = search_signature
            st.session_state.viewer_page = None

    results = st.session_state.get("search_results", []) if should_search else []

    # --------------------------------------------------------
    # MAIN BODY — left browser/results + right PDF reader
    # --------------------------------------------------------
    left_col, right_col = st.columns([1.12, 1.0], gap="large")

    with left_col:
        # Browse section title follows the reference image.
        browse_label = {
            "Category": "Browse by Category",
            "Device": "Browse by Device",
            "Topic": "Browse by Topic",
        }.get(tile_type, "Browse by Category")
        st.markdown(
            f'<div class="browse-heading"><span>{browse_label}</span><span class="browse-viewall">View All</span></div>',
            unsafe_allow_html=True,
        )

        active_tiles = tile_data.get(tile_type, [])
        show_all = st.session_state.get("knowledge_tile_show_all", False)
        visible_tiles = active_tiles if show_all else active_tiles[:6]

        if visible_tiles:
            tile_cols = st.columns(6, gap="small")
            tile_icons = {
                "Category": ["▤", "♧", "✚", "⬢", "▤", "⌁"],
                "Device": ["♧", "⬡", "▰", "◉", "▣", "▥"],
                "Topic": ["⇩", "⚙", "☷", "⬢", "●", "⚠"],
            }
            icons = tile_icons.get(tile_type, [])

            for idx, (label, count) in enumerate(visible_tiles):
                with tile_cols[idx % 6]:
                    selected = selected_tile == label
                    icon = icons[idx % len(icons)] if icons else "▤"
                    if st.button(
                        f"{icon}\n{label}\n{count} PDFs",
                        key=f"knowledge_tile_{tile_type}_{idx}_{label}",
                        use_container_width=True,
                        type="primary" if selected else "secondary",
                    ):
                        st.session_state.knowledge_tile_value = label
                        st.session_state.knowledge_tile_show_all = False
                        if tile_type == "Category":
                            st.session_state.search_category = label
                            st.session_state.search_query = ""
                        else:
                            st.session_state.search_query = label
                        st.session_state.selected_result_id = None
                        st.session_state.search_signature = None
                        st.rerun()

            if len(active_tiles) > 6:
                if st.button(
                    "Show fewer" if show_all else "›",
                    key=f"view_all_{tile_type}",
                    help="Show all available tiles" if not show_all else "Show fewer tiles",
                ):
                    st.session_state.knowledge_tile_show_all = not show_all
                    st.rerun()
        else:
            st.caption(f"No {tile_type.lower()} tiles were detected in the indexed PDFs yet.")

        # Search results stay directly below the browse tiles.
        if results:
            st.markdown(
                f'<div class="search-results-heading"><span>Search Results</span><span class="sort-chip">Sort by: Relevance ˅</span></div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                f'<div class="search-count">{len(results)} result(s) for “{active_query or selected_tile or category}” — Showing the highest match first</div>',
                unsafe_allow_html=True,
            )

            for i, result in enumerate(results):
                is_best = i == 0
                key_prefix = "best" if is_best else str(i)
                is_selected = result["id"] == st.session_state.get("selected_result_id")
                with st.container(key=f"search_result_{key_prefix}", border=True):
                    score_pct = min(99, max(1, round(result["score"] * 100))) if result["score"] else 0
                    label = "Highest Match" if is_best else "Search Result"
                    if is_selected:
                        label = "Selected Result"
                    st.markdown(
                        f"""
                        <div class='result-label'>{label}<span class='result-score-pill'>{score_pct}%</span></div>
                        <div class='result-filename'>{result['filename']}</div>
                        <div class='result-page'>Page {result['page_number']}</div>
                        <div class='result-snippet'>{result['snippet']}</div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if st.button("Open Result", key=f"open_result_{result['id']}", use_container_width=True):
                        st.session_state.selected_result_id = result["id"]
                        st.session_state.viewer_page = int(result["page_number"])
                        st.rerun()
        elif should_search:
            st.warning("No matching PDF was found. Try different keywords or upload another document.")

    with right_col:
        if results:
            best = results[0]
            selected_result_id = st.session_state.get("selected_result_id")
            selected_matches = [r for r in results if r["id"] == selected_result_id]
            selected_result = selected_matches[0] if selected_matches else best

            if selected_result_id != selected_result["id"]:
                st.session_state.selected_result_id = selected_result["id"]
                st.session_state.viewer_page = int(selected_result["page_number"])

            selected_doc = get_document_by_id(selected_result["document_id"])
            total_pages = int(selected_doc["page_count"] or 0) if selected_doc else 0
            if st.session_state.get("viewer_page") is None:
                st.session_state.viewer_page = int(selected_result["page_number"])
            viewer_page = max(1, min(int(st.session_state.viewer_page), max(total_pages, 1)))

            st.markdown(
                f"""
                <div class='reader-toolbar'>
                    <div class='reader-title'>▤ {selected_result['filename']}</div>
                    <div class='reader-meta'>Page {viewer_page} of {total_pages} · Search: “{active_query or selected_tile or ''}”</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            zoom = st.selectbox(
                "Zoom",
                [100, 125, 150, 175, 200],
                index=0,
                format_func=lambda x: f"{x}%",
                label_visibility="collapsed",
                key="pdf_zoom",
            )

            render_scale = zoom / 100 * 1.05
            image = render_pdf_page_highlighted(
                selected_result["stored_path"],
                viewer_page,
                active_query or selected_tile or "",
                scale=render_scale,
            )

            if image:
                display_width = {100: 760, 125: 860, 150: 960, 175: 1080, 200: 1200}.get(int(zoom), 760)
                with st.container(key="pdf_viewer_shell"):
                    with st.container(key="pdf_prev_wrap"):
                        if st.button(
                            "‹",
                            disabled=(viewer_page <= 1),
                            key="pdf_prev",
                            help="Previous page",
                        ):
                            st.session_state.viewer_page = max(1, viewer_page - 1)
                            st.rerun()
                    with st.container(key="pdf_next_wrap"):
                        if st.button(
                            "›",
                            disabled=(viewer_page >= total_pages),
                            key="pdf_next",
                            help="Next page",
                        ):
                            st.session_state.viewer_page = min(total_pages, viewer_page + 1)
                            st.rerun()
                    st.image(image, width=display_width)
            else:
                st.error("Unable to render this PDF page.")

            matches = find_document_matches(
                selected_result["stored_path"],
                active_query or selected_tile or "",
                limit=8,
            )
            if matches:
                with st.expander("Matches in this PDF", expanded=False):
                    for match in matches:
                        if st.button(
                            f"Page {match['page']}",
                            key=f"jump_match_{selected_result['id']}_{match['page']}",
                            use_container_width=True,
                        ):
                            st.session_state.viewer_page = match["page"]
                            st.rerun()
                        st.markdown(
                            f"<div class='match-item'><div class='match-item-page'>Page {match['page']} · {match['score']} term match(es)</div><div class='match-item-text'>{match['snippet']}</div></div>",
                            unsafe_allow_html=True,
                        )
        else:
            st.markdown(
                """
                <div class="welcome-card">
                    <div class="welcome-icon">⌕</div>
                    <div class="welcome-title">Search your knowledge base</div>
                    <div class="welcome-text">Search for a process, checklist, policy, troubleshooting step, or any information contained in your uploaded PDFs.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

# ============================================================

st.markdown('<div class="bottom-nav-spacer"></div>', unsafe_allow_html=True)
nav_left, nav_center, nav_right = st.columns([1, 2, 1])
with nav_center:
    if st.button("⌕  Search Knowledge Base", use_container_width=True):
        st.session_state.page = "Search"; st.rerun()

st.markdown('<div class="bottom-nav-label">Knowledge Base · PDF Search</div>', unsafe_allow_html=True)
