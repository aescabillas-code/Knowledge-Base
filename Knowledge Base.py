import os
import re
import io
import sqlite3
import hashlib
import hmac
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

try:
    from itsdangerous import URLSafeTimedSerializer
except Exception:
    URLSafeTimedSerializer = None


# ============================================================
# HPE KNOWLEDGE BASE
# Combined:
# 1) Image-matched UI
# 2) PDF upload + extraction + TF-IDF search
# 3) Searchable PDF viewer
# 4) Admin document management
# 5) One-time access-code gate with signed browser token
# 6) Mock/public HPE content so the UI is populated immediately
#
# IMPORTANT:
# The seeded content is concise paraphrased information based on
# public HPE pages. It is not presented as proprietary/internal HPE SOP.
#
# Public references used for the seeded mock data:
# - https://www.hpe.com/us/en/products/compute/proliant.html
# - https://www.hpe.com/us/en/collaterals/collateral.c04154343.html
# - https://developer.hpe.com/platform/hpe-greenlake/home/
# - https://www.hpe.com/us/en/aruba-central.html.html
# - https://arubanetworking.hpe.com/techdocs/central/latest/content/home.htm
# - https://arubanetworking.hpe.com/techdocs/central/2.5.8/content/nms/landing-pages/switches.htm
# - https://www.hpe.com/us/en/networking.html
# - https://www.hpe.com/sg/en/products/networking/hpe-aruba-networking-services.html
# ============================================================


APP_NAME = "HPE Knowledge Base"
DATA_DIR = Path("knowledge_base_data")
PDF_DIR = DATA_DIR / "pdfs"
DB_PATH = DATA_DIR / "knowledge_base.db"

DATA_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CONFIG / SECRETS
# ============================================================

def get_secret(name, default=""):
    try:
        value = st.secrets.get(name)
        if value is not None:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


ACCESS_CODE = get_secret("ACCESS_CODE", "HPE-DEMO-2026")
ADMIN_PIN = get_secret("ADMIN_PIN", "2468")
TOKEN_SECRET = get_secret("TOKEN_SECRET", "")

if not TOKEN_SECRET:
    TOKEN_SECRET = hashlib.sha256(
        f"{os.getcwd()}::{os.getenv('HOSTNAME', 'hpe-kb')}".encode()
    ).hexdigest()


# ============================================================
# CSS — MATCHED TO THE UPLOADED REFERENCE
# ============================================================

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root{
    --ink:#102238;
    --ink2:#172938;
    --muted:#6d7884;
    --line:#e5e9ed;
    --canvas:#f5f7f9;
    --white:#ffffff;
    --nav:#00363b;
    --nav-deep:#002c31;
    --green:#01a982;
    --green2:#009f84;
}

html, body, [class*="css"]{
    font-family:"Inter",Arial,sans-serif;
}

.stApp{
    background:var(--canvas);
    color:var(--ink);
}

#MainMenu, footer, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"], [data-testid="stMainMenu"]{
    display:none !important;
    visibility:hidden !important;
}

[data-testid="stHeader"]{
    background:transparent !important;
    height:0 !important;
    min-height:0 !important;
}

.block-container{
    max-width:none !important;
    padding:0 18px 30px 230px !important;
}

/* ---------------- LEFT NAV ---------------- */

[data-testid="stSidebar"]{
    width:212px !important;
    min-width:212px !important;
    background:linear-gradient(180deg,#003d42 0%,#003037 100%) !important;
    border-right:0 !important;
}

[data-testid="stSidebar"] > div:first-child,
[data-testid="stSidebarContent"]{
    padding:0 !important;
}

[data-testid="stSidebarHeader"]{
    height:0 !important;
    min-height:0 !important;
}

[data-testid="stSidebarCollapseButton"],
[data-testid="stSidebarCollapsedControl"]{
    display:none !important;
}

.sidebar-logo{
    height:62px;
    background:#fafbfd;
    color:#111b25;
    display:flex;
    align-items:center;
    padding-left:21px;
    font-size:34px;
    font-weight:800;
    letter-spacing:-2.8px;
    border-bottom:1px solid #e5e9ec;
}

.sidebar-logo .e{
    color:#01a982;
    margin-left:1px;
}

.sidebar-space{
    height:29px;
}

[data-testid="stSidebar"] div.stButton{
    margin:0 11px 3px !important;
}

[data-testid="stSidebar"] div.stButton > button{
    width:100% !important;
    min-height:41px !important;
    border:0 !important;
    border-radius:8px !important;
    background:transparent !important;
    color:#f1f8f8 !important;
    box-shadow:none !important;
    text-align:left !important;
    padding:0 13px !important;
    font-size:14px !important;
    font-weight:500 !important;
}

[data-testid="stSidebar"] div.stButton > button:hover{
    background:rgba(1,169,130,.14) !important;
}

[data-testid="stSidebar"] div.stButton > button p{
    font-size:14px !important;
    color:inherit !important;
}

.sidebar-footer{
    position:fixed;
    left:23px;
    bottom:17px;
    color:white;
    pointer-events:none;
}
.sidebar-footer .big{
    font-size:21px;
    font-weight:700;
}
.sidebar-footer .small{
    font-size:12px;
}
.sidebar-footer .version{
    font-size:10px;
    opacity:.6;
    margin-top:5px;
}

/* ---------------- HEADER ---------------- */

.topbar{
    height:62px;
    display:flex;
    align-items:center;
}

.page-heading{
    font-size:24px;
    line-height:1;
    font-weight:750;
    letter-spacing:-.6px;
    color:#112438;
    padding-top:2px;
}

.profile-name{
    font-size:12px;
    font-weight:650;
    color:#182738;
}
.profile-role{
    font-size:11px;
    color:#6d7883;
    margin-top:2px;
}
.avatar{
    width:39px;
    height:39px;
    border-radius:50%;
    display:flex;
    align-items:center;
    justify-content:center;
    background:linear-gradient(145deg,#efd0c1,#fff);
    border:2px solid #fff;
    box-shadow:0 1px 5px rgba(0,0,0,.15);
    font-size:18px;
}

/* ---------------- HERO ---------------- */

.hero{
    min-height:214px;
    border-radius:9px;
    overflow:hidden;
    padding:31px 38px 18px;
    position:relative;
    color:white;
    background:
      radial-gradient(circle at 82% 30%,rgba(20,230,207,.25),transparent 23%),
      radial-gradient(circle at 98% 65%,rgba(16,215,192,.15),transparent 28%),
      linear-gradient(108deg,#0c2027 0%,#0a4b52 48%,#007e78 100%);
    box-shadow:0 3px 13px rgba(18,44,53,.08);
}

.hero::before{
    content:"";
    position:absolute;
    right:-20px;
    top:-20px;
    width:470px;
    height:260px;
    opacity:.28;
    background:
      linear-gradient(155deg,transparent 46%,#1de2c8 46.7%,transparent 48%),
      linear-gradient(165deg,transparent 59%,#1de2c8 59.7%,transparent 61%),
      linear-gradient(145deg,transparent 69%,#1de2c8 69.7%,transparent 71%);
    transform:skewX(-8deg);
}

.hero h1{
    margin:0;
    font-size:36px;
    font-weight:750;
    letter-spacing:-1.2px;
    line-height:1.1;
    position:relative;
    z-index:2;
}
.hero-sub{
    margin-top:8px;
    font-size:16px;
    position:relative;
    z-index:2;
}

.popular-row{
    display:flex;
    align-items:center;
    gap:8px;
    flex-wrap:wrap;
    margin-top:10px;
    position:relative;
    z-index:2;
}
.popular-text{
    font-size:12px;
}
.popular-pill{
    background:rgba(8,21,28,.38);
    border:1px solid rgba(255,255,255,.12);
    border-radius:15px;
    padding:5px 12px;
    font-size:11px;
}

/* ---------------- INPUTS ---------------- */

div[data-testid="stTextInput"] input{
    height:40px !important;
    border-radius:8px !important;
    border:1px solid #dfe4e8 !important;
    background:white !important;
    color:#142435 !important;
    font-size:13px !important;
}

.hero-input div[data-testid="stTextInput"] input{
    height:51px !important;
    border-radius:10px !important;
    font-size:14px !important;
}

div[data-testid="stTextInput"] label{
    display:none !important;
}

/* ---------------- BUTTONS ---------------- */

div.stButton > button,
div.stFormSubmitButton > button{
    border-radius:8px !important;
    border:1px solid #dce2e6 !important;
    background:#fff !important;
    color:#172635 !important;
    font-weight:600 !important;
    min-height:39px;
}

div.stButton > button:hover,
div.stFormSubmitButton > button:hover{
    border-color:#01a982 !important;
    color:#007e69 !important;
}

button[kind="primary"]{
    background:#01a982 !important;
    border-color:#01a982 !important;
    color:white !important;
}

/* ---------------- METRICS ---------------- */

.metrics{
    margin-top:15px;
}

.metric{
    height:84px;
    border:1px solid #e7ebef;
    background:white;
    border-radius:10px;
    display:flex;
    align-items:center;
    padding:0 14px;
    box-shadow:0 2px 8px rgba(20,40,55,.025);
}

.metric-icon{
    width:48px;
    height:48px;
    border-radius:15px;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:23px;
    flex:0 0 auto;
}
.metric-body{
    padding-left:13px;
    flex:1;
}
.metric-number{
    font-size:23px;
    line-height:1;
    font-weight:750;
}
.metric-label{
    font-size:12px;
    margin-top:5px;
    color:#263342;
}
.metric-arrow{
    font-size:22px;
    color:#1d2b36;
}

/* ---------------- CARDS ---------------- */

.white-card{
    background:white;
    border:1px solid #e5e9ed;
    border-radius:11px;
    box-shadow:0 2px 9px rgba(20,40,55,.025);
    padding:13px;
}

.section-title{
    font-size:16px;
    font-weight:750;
    color:#102238;
}
.section-head{
    display:flex;
    justify-content:space-between;
    align-items:center;
    padding:0 2px 10px;
}
.section-link{
    font-size:12px;
}

/* ---------------- CATEGORIES ---------------- */

.category-tile{
    min-height:69px;
    border:1px solid #e6eaee;
    border-radius:9px;
    display:flex;
    align-items:center;
    padding:9px 11px;
    background:#fff;
}

.category-icon{
    width:39px;
    height:39px;
    border-radius:50%;
    display:flex;
    justify-content:center;
    align-items:center;
    margin-right:10px;
    font-size:19px;
}

.category-name{
    font-size:12px;
    font-weight:650;
}
.category-count{
    font-size:10px;
    color:#78828d;
    margin-top:4px;
}
.category-arrow{
    margin-left:auto;
    font-size:19px;
}

/* ---------------- RECENT DOCS ---------------- */

.recent-row{
    height:58px;
    display:flex;
    align-items:center;
    border-bottom:1px solid #edf0f2;
}
.recent-row:last-child{
    border-bottom:0;
}
.pdf-icon{
    width:33px;
    height:37px;
    border-radius:7px;
    background:#fff0f1;
    color:#db2734;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:10px;
    font-weight:800;
    margin-right:11px;
}
.recent-title{
    font-size:12px;
    font-weight:650;
}
.recent-meta{
    color:#78838e;
    font-size:10px;
    margin-top:3px;
}
.recent-date{
    margin-left:auto;
    color:#707b85;
    font-size:10px;
    white-space:nowrap;
}

/* ---------------- FEATURED PDF MOCK VIEW ---------------- */

.viewer{
    height:257px;
    background:#20262b;
    border-radius:6px;
    padding:13px;
    display:flex;
    justify-content:center;
    overflow:hidden;
}
.paper{
    width:83%;
    height:100%;
    background:#fff;
    box-shadow:0 3px 14px rgba(0,0,0,.24);
    padding:22px 26px;
    color:#142536;
    overflow:hidden;
}
.paper-logo{
    font-size:22px;
    font-weight:850;
    letter-spacing:-2px;
}
.paper-logo span{
    color:#01a982;
}
.paper-title{
    font-size:19px;
    font-weight:750;
    margin-top:12px;
}
.paper-rule{
    height:3px;
    width:116px;
    background:#01a982;
    margin:13px 0;
}
.paper-building{
    height:70px;
    margin:4px -26px 10px;
    background:
      linear-gradient(135deg,transparent 35%,rgba(50,150,210,.25) 35%,rgba(50,150,210,.09) 58%,transparent 59%),
      linear-gradient(160deg,#d9eef8,#f9fcff 70%);
}
.paper-small{
    font-size:9px;
    color:#697681;
}

/* ---------------- ACCESS GATE ---------------- */

.access-wrap{
    max-width:520px;
    margin:90px auto 0;
    background:white;
    border:1px solid #e4e9ed;
    border-radius:13px;
    padding:38px;
    text-align:center;
    box-shadow:0 14px 42px rgba(0,44,54,.08);
}
.access-logo{
    font-size:48px;
    line-height:1;
    font-weight:850;
    letter-spacing:-4px;
}
.access-logo span{
    color:#01a982;
}
.access-title{
    margin-top:14px;
    font-size:24px;
    font-weight:750;
}
.access-sub{
    color:#6f7b86;
    font-size:13px;
    margin-top:7px;
}

/* ---------------- ADMIN ---------------- */

.admin-bar{
    border-radius:9px;
    background:linear-gradient(90deg,#073d42,#007f72);
    color:#fff;
    padding:17px 21px;
    margin-bottom:15px;
}
.admin-title{
    font-size:20px;
    font-weight:750;
}
.admin-sub{
    font-size:12px;
    opacity:.9;
    margin-top:4px;
}

.admin-card{
    max-width:520px;
    margin:60px auto;
    text-align:center;
    background:#fff;
    border:1px solid #e4e9ed;
    border-radius:13px;
    padding:32px;
}

.admin-icon{
    font-size:38px;
}
.admin-title-big{
    font-size:24px;
    font-weight:750;
    margin-top:7px;
}
.admin-subtitle{
    color:#6f7b86;
    font-size:13px;
    margin-top:5px;
}

/* ---------------- SEARCH RESULT / READER ---------------- */

.search-count{
    color:#687b87;
    font-size:11px;
    margin:2px 0 8px;
}

.result-label{
    color:#0561a0;
    font-weight:700;
    font-size:11px;
    line-height:1.2;
}
.result-filename{
    margin-top:3px;
    color:#0561a0;
    font-weight:700;
    font-size:12px;
    line-height:1.35;
    word-break:break-word;
}
.result-score{
    margin-top:3px;
    color:#687b87;
    font-size:10px;
}
.result-snippet{
    color:#536b78;
    font-size:10px;
    line-height:1.4;
    margin-top:5px;
    display:-webkit-box;
    -webkit-line-clamp:3;
    -webkit-box-orient:vertical;
    overflow:hidden;
}
.result-score-pill{
    float:right;
    background:#e7f8f1;
    color:#087c63;
    border-radius:10px;
    padding:2px 6px;
    font-size:9px;
    font-weight:700;
}
.result-page{
    color:#687b87;
    font-size:10px;
    margin-top:3px;
}

.reader-toolbar{
    background:#ffffff;
    border:1px solid var(--line);
    border-radius:8px;
    padding:7px 10px;
    margin-bottom:8px;
}
.reader-title{
    font-size:15px;
    font-weight:700;
    color:var(--ink);
    line-height:1.3;
    word-break:break-word;
}
.reader-meta{
    font-size:11px;
    color:var(--muted);
    margin-top:2px;
}
.reader-match{
    background:#e7f8f1;
    border:1px solid #9bdcc8;
    color:#087c63;
    border-radius:5px;
    padding:5px 8px;
    font-size:10px;
    font-weight:700;
    display:inline-block;
    margin-top:5px;
}

.reader-page-indicator{
    text-align:right;
    color:#687b87;
    font-size:11px;
    line-height:30px;
    padding-right:4px;
}

.match-panel{
    background:#ffffff;
    border:1px solid var(--line);
    border-radius:8px;
    padding:10px;
    margin-top:10px;
}
.match-panel-title{
    font-size:13px;
    font-weight:700;
    color:var(--ink);
    margin-bottom:7px;
}
.match-item{
    background:#f7fafb;
    border:1px solid #e1e8ec;
    border-radius:6px;
    padding:7px 8px;
    margin-bottom:6px;
}
.match-item-page{
    color:#0561a0;
    font-size:10px;
    font-weight:700;
}
.match-item-text{
    color:#385362;
    font-size:10px;
    line-height:1.35;
    margin-top:2px;
}

/* ---------------- PDF VIEWER FLOATING ARROWS ---------------- */

.st-key-pdf_viewer_shell{
    position:relative !important;
    overflow:visible !important;
    padding:0 !important;
}

.st-key-pdf_prev_wrap,
.st-key-pdf_next_wrap{
    position:absolute !important;
    top:50% !important;
    transform:translateY(-50%) !important;
    z-index:20 !important;
    width:42px !important;
}

.st-key-pdf_prev_wrap{
    left:10px !important;
}

.st-key-pdf_next_wrap{
    right:10px !important;
}

.st-key-pdf_prev_wrap button,
.st-key-pdf_next_wrap button{
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
}

.st-key-pdf_prev_wrap button:hover:not(:disabled),
.st-key-pdf_next_wrap button:hover:not(:disabled){
    transform:scale(1.06) !important;
}

.st-key-pdf_prev_wrap button:disabled,
.st-key-pdf_next_wrap button:disabled{
    opacity:.35 !important;
}

/* ---------------- WELCOME ---------------- */

.welcome-card{
    margin:42px auto;
    max-width:720px;
    text-align:center;
    background:#fff;
    border:1px solid var(--line);
    border-radius:14px;
    padding:42px;
    box-shadow:0 5px 18px rgba(12,54,70,.05);
}
.welcome-icon{
    font-size:42px;
    color:var(--green);
}
.welcome-title{
    font-size:25px;
    font-weight:700;
    color:var(--ink);
    margin-top:8px;
}
.welcome-text{
    color:var(--muted);
    max-width:560px;
    margin:10px auto;
    line-height:1.6;
    font-size:14px;
}
.welcome-stats{
    display:flex;
    justify-content:center;
    gap:35px;
    color:#57707e;
    margin-top:18px;
    font-size:12px;
}

/* ---------------- RESPONSIVE ---------------- */

@media(max-width:1050px){
    .block-container{
        padding-left:225px !important;
    }
}

@media(max-width:720px){
    [data-testid="stSidebar"]{
        width:70px !important;
        min-width:70px !important;
    }

    .block-container{
        padding-left:82px !important;
        padding-right:9px !important;
    }

    .sidebar-logo{
        font-size:25px;
        padding-left:13px;
    }

    [data-testid="stSidebar"] div.stButton > button{
        font-size:0 !important;
        text-align:center !important;
        padding:0 !important;
    }

    .sidebar-footer{
        display:none;
    }

    .hero h1{
        font-size:28px;
    }
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
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def init_db():
    conn = db()

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            stored_path TEXT NOT NULL DEFAULT '',
            file_hash TEXT UNIQUE NOT NULL,
            category TEXT DEFAULT 'General',
            title TEXT DEFAULT '',
            description TEXT DEFAULT '',
            keywords TEXT DEFAULT '',
            source_url TEXT DEFAULT '',
            source_type TEXT DEFAULT 'PDF',
            page_count INTEGER DEFAULT 0,
            file_size INTEGER DEFAULT 0,
            uploaded_at TEXT NOT NULL,
            indexed_at TEXT,
            status TEXT DEFAULT 'Indexed',
            featured INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS chunks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER NOT NULL,
            page_number INTEGER NOT NULL,
            chunk_index INTEGER NOT NULL,
            text TEXT NOT NULL,
            FOREIGN KEY(document_id) REFERENCES documents(id)
        );

        CREATE INDEX IF NOT EXISTS idx_chunks_document_page
        ON chunks(document_id, page_number);

        CREATE INDEX IF NOT EXISTS idx_documents_category
        ON documents(category);
        """
    )

    conn.commit()
    conn.close()


init_db()


# ============================================================
# SESSION STATE
# ============================================================

defaults = {
    "access_authorized": False,
    "admin_authenticated": False,
    "page": "Home",
    "search_query": "",
    "search_results": [],
    "search_signature": None,
    "selected_result_id": None,
    "viewer_page": 1,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# PUBLIC HPE MOCK DATA
# ============================================================
# Concise paraphrases, not copied document text.
# Each record includes a public HPE URL so the UI can point users
# to the authoritative source.
# ============================================================

MOCK_DOCS = [
    {
        "filename": "HPE ProLiant Compute Overview",
        "title": "HPE ProLiant Compute Overview",
        "category": "Compute",
        "description": (
            "HPE ProLiant Compute is HPE's server portfolio for modern hybrid "
            "workloads. HPE describes the portfolio as supporting security, "
            "performance, automation and energy-efficient operations across "
            "workloads ranging from enterprise applications to AI."
        ),
        "keywords": (
            "ProLiant, compute, servers, rack, tower, edge, AI, Gen11, Gen12, "
            "HPE Compute, hybrid cloud"
        ),
        "source_url": "https://www.hpe.com/us/en/products/compute/proliant.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 1,
        "text": (
            "HPE ProLiant Compute is a portfolio of HPE servers designed for "
            "modern hybrid environments. Public HPE product information describes "
            "security-first design, performance and efficiency, and AI-driven "
            "automation as key capabilities. The portfolio includes rack, tower, "
            "edge and AI-oriented systems. HPE also describes differences between "
            "Gen11 and Gen12 systems, including newer processors, management "
            "capabilities, memory and storage options, and support for demanding "
            "workloads depending on the model."
        ),
    },
    {
        "filename": "HPE Integrated Lights-Out (iLO) Quick Reference",
        "title": "HPE Integrated Lights-Out (iLO) Quick Reference",
        "category": "Management",
        "description": (
            "Public HPE QuickSpecs describe iLO as embedded server technology "
            "used for server setup, health monitoring, and power and thermal control."
        ),
        "keywords": (
            "iLO, Integrated Lights-Out, server management, health monitoring, "
            "power, thermal, ProLiant, iLO 5, iLO 6, iLO 7"
        ),
        "source_url": "https://www.hpe.com/us/en/collaterals/collateral.c04154343.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE Integrated Lights-Out, commonly called iLO, is embedded technology "
            "in HPE servers. HPE describes iLO as a foundation for server intelligence "
            "and lists capabilities that help with server setup, health monitoring, "
            "and power and thermal control. The iLO ASIC and firmware vary by server "
            "generation. HPE's current public QuickSpecs identify iLO 7 with Gen12 "
            "systems and also document iLO 6 and iLO 5 for applicable generations. "
            "The exact iLO version should therefore be verified against the specific "
            "server model and generation."
        ),
    },
    {
        "filename": "HPE GreenLake Platform Overview",
        "title": "HPE GreenLake Platform Overview",
        "category": "HPE GreenLake",
        "description": (
            "HPE describes GreenLake as an edge-to-cloud platform intended to "
            "provide a cloud experience across distributed applications and data."
        ),
        "keywords": (
            "HPE GreenLake, edge to cloud, hybrid cloud, cloud experience, "
            "self service, APIs, platform"
        ),
        "source_url": "https://developer.hpe.com/platform/hpe-greenlake/home/",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE GreenLake is described by HPE as an edge-to-cloud platform that "
            "brings a cloud experience to distributed applications and data. HPE "
            "also describes scalable consumption and self-service experiences for "
            "its services. The HPE GreenLake developer portal provides public API "
            "documentation and technical resources for customers and partners. "
            "When troubleshooting a GreenLake issue, identify the specific service, "
            "workspace, account context and role before using a service-specific "
            "procedure."
        ),
    },
    {
        "filename": "HPE Aruba Networking Central Overview",
        "title": "HPE Aruba Networking Central Overview",
        "category": "Networking",
        "description": (
            "HPE Aruba Networking Central is described in public HPE documentation "
            "as a cloud-managed networking platform for wired, wireless, WAN and VPN services."
        ),
        "keywords": (
            "Aruba Central, HPE Aruba Networking Central, cloud managed, wired, "
            "wireless, WAN, VPN, monitoring, troubleshooting"
        ),
        "source_url": "https://www.hpe.com/us/en/aruba-central.html.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE Aruba Networking Central is a cloud-managed networking platform "
            "for wired, wireless, WAN and VPN services. HPE documentation describes "
            "Central as providing unified visibility and management, automation, "
            "analytics, monitoring, reporting and troubleshooting capabilities. "
            "Public Aruba documentation also describes Central as a management and "
            "orchestration console across supported networking domains. When handling "
            "a Central support question, identify the affected site, device type, "
            "service and observed alert or health metric before selecting a procedure."
        ),
    },
    {
        "filename": "HPE Aruba Networking Switches Reference",
        "title": "HPE Aruba Networking Switches Reference",
        "category": "Networking",
        "description": (
            "Public HPE Aruba documentation describes Aruba switches across edge "
            "access and data-center use cases, with AOS-S and AOS-CX families."
        ),
        "keywords": (
            "Aruba switches, AOS-S, AOS-CX, switching, access, data center, "
            "cloud native, network access"
        ),
        "source_url": "https://arubanetworking.hpe.com/techdocs/central/2.5.8/content/nms/landing-pages/switches.htm",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE Aruba Networking public documentation describes switches that "
            "support secure, role-based network access for wired users and devices. "
            "The portfolio includes switches from edge access through data-center "
            "environments. Aruba documentation references AOS-S and AOS-CX switch "
            "families in Central-related workflows. For a support case, confirm "
            "the switch family, operating system, site context and intended task "
            "before applying device-specific instructions."
        ),
    },
    {
        "filename": "HPE Aruba Networking Portfolio Overview",
        "title": "HPE Aruba Networking Portfolio Overview",
        "category": "Networking",
        "description": (
            "HPE's networking portfolio includes Central, Wi-Fi, CX switches, "
            "SSE, SD-WAN, private 5G and related networking services."
        ),
        "keywords": (
            "Aruba Networking, Wi-Fi, CX switches, SSE, SD-WAN, private 5G, "
            "networking portfolio, edge to cloud"
        ),
        "source_url": "https://www.hpe.com/us/en/networking.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE's public networking portfolio includes HPE Aruba Networking "
            "Central, Wi-Fi solutions, CX switches, SSE, EdgeConnect SD-WAN, "
            "private 5G and other networking services. HPE describes the portfolio "
            "as supporting secure connectivity, cloud-native management, automation "
            "and AI-related operational capabilities. The correct product family "
            "should be identified before selecting a troubleshooting or deployment "
            "reference."
        ),
    },
    {
        "filename": "HPE Aruba Networking Services Lifecycle Reference",
        "title": "HPE Aruba Networking Services Lifecycle Reference",
        "category": "Support",
        "description": (
            "HPE describes networking services spanning planning and design, "
            "deployment and migration, ongoing operations, support and optimization."
        ),
        "keywords": (
            "Aruba services, support, training, professional services, deployment, "
            "migration, optimization, lifecycle"
        ),
        "source_url": "https://www.hpe.com/sg/en/products/networking/hpe-aruba-networking-services.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE Aruba Networking services are described as supporting the network "
            "lifecycle from planning and design through deployment, ongoing operation "
            "and optimization. Public HPE information also references support, "
            "education and professional services. For a support workflow, determine "
            "whether the request concerns planning, deployment, operations, training "
            "or optimization before selecting the relevant service path."
        ),
    },
    {
        "filename": "HPE GreenLake Backup and Recovery Overview",
        "title": "HPE GreenLake Backup and Recovery Overview",
        "category": "Data Protection",
        "description": (
            "HPE's public getting-started guide describes GreenLake for Backup and "
            "Recovery as a data protection service within the HPE GreenLake platform."
        ),
        "keywords": (
            "GreenLake Backup Recovery, backup, recovery, VMware, SQL Server, "
            "array volumes, Data Services Cloud Console"
        ),
        "source_url": "https://support.hpe.com/hpesc/public/docDisplay?docId=sd00003454en_us&docLocale=en_US&page=GUID-76A1AF84-7118-45D9-B2E3-F0166D955259.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE GreenLake for Backup and Recovery is described in HPE's public "
            "getting-started documentation as a data protection service within "
            "Data Services Cloud Console on the HPE GreenLake platform. The service "
            "supports backup and restore scenarios for resources including VMware "
            "virtual machines and datastores, Microsoft SQL Server databases and "
            "instances, and HPE array volumes. HPE notes that users need an applicable "
            "subscription and an HPE GreenLake Cloud Portal account with the required roles."
        ),
    },
    {
        "filename": "HPE Compute Management Quick Reference",
        "title": "HPE Compute Management Quick Reference",
        "category": "Compute",
        "description": (
            "Public HPE ProLiant information identifies HPE iLO and HPE Compute Ops "
            "Management among management technologies used with supported systems."
        ),
        "keywords": (
            "Compute Ops Management, iLO, ProLiant, management, automation, "
            "server operations"
        ),
        "source_url": "https://www.hpe.com/us/en/products/compute/proliant.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": (
            "HPE public ProLiant information identifies HPE iLO and HPE Compute Ops "
            "Management as management technologies associated with supported systems. "
            "The exact management capability depends on the server model and generation. "
            "For a case involving server management, confirm the model, generation, "
            "iLO version and management service before selecting the relevant procedure."
        ),
    },

    {
        "filename": "HPE ProLiant Server Family Guide",
        "title": "HPE ProLiant Server Family Guide",
        "category": "Product",
        "description": "A demo reference for identifying rack, tower, edge and AI-oriented HPE ProLiant systems before troubleshooting or configuration.",
        "keywords": "ProLiant rack tower edge AI server model identification product family",
        "source_url": "https://www.hpe.com/us/en/products/compute/proliant.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Use the HPE ProLiant portfolio to identify the appropriate server family before selecting a support workflow. Public HPE information groups systems around rack, tower, edge and AI-oriented use cases. For a case, capture the exact product name, generation and configuration because capabilities differ by model."
    },
    {
        "filename": "HPE ProLiant Gen11 and Gen12 Comparison Notes",
        "title": "HPE ProLiant Gen11 and Gen12 Comparison Notes",
        "category": "Product",
        "description": "Demo notes for comparing supported ProLiant generations and verifying management, processor, memory and storage capabilities.",
        "keywords": "ProLiant Gen11 Gen12 comparison processor memory storage iLO",
        "source_url": "https://www.hpe.com/us/en/products/compute/proliant.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE public product information describes differences across ProLiant generations. When comparing Gen11 and Gen12, verify the exact server model, processor options, memory and storage configuration, and supported management technologies rather than assuming that a capability applies to every model."
    },
    {
        "filename": "HPE iLO Server Health Monitoring Guide",
        "title": "HPE iLO Server Health Monitoring Guide",
        "category": "Management",
        "description": "Quick demo reference for using iLO concepts when reviewing server health, setup and power or thermal information.",
        "keywords": "iLO health monitoring server setup power thermal management",
        "source_url": "https://www.hpe.com/us/en/collaterals/collateral.c04154343.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE iLO is embedded server technology used for server management. Public HPE QuickSpecs describe capabilities related to server setup, health monitoring and power or thermal control. Before using an iLO procedure, verify the server model and iLO generation because features can vary by platform."
    },
    {
        "filename": "HPE Compute Ops Management Overview",
        "title": "HPE Compute Ops Management Overview",
        "category": "Management",
        "description": "Demo reference for understanding the role of HPE Compute Ops Management alongside supported ProLiant systems.",
        "keywords": "Compute Ops Management ProLiant server operations automation cloud management",
        "source_url": "https://www.hpe.com/us/en/products/compute/proliant.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE public ProLiant information identifies Compute Ops Management as a management technology for supported systems. A support workflow should first establish the server model, generation, management service and customer environment before selecting an operations or automation procedure."
    },
    {
        "filename": "HPE GreenLake Workspace and Access Basics",
        "title": "HPE GreenLake Workspace and Access Basics",
        "category": "HPE GreenLake",
        "description": "Demo troubleshooting reference for identifying workspace, account and role context in HPE GreenLake.",
        "keywords": "GreenLake workspace account role access permissions cloud portal",
        "source_url": "https://developer.hpe.com/greenlake/hpe-greenlake-platform/home/",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "For an HPE GreenLake access issue, identify the affected workspace, account context and user role. Public HPE developer documentation describes a unified platform with service APIs and role-based access concepts. Exact permissions depend on the service and account configuration."
    },
    {
        "filename": "HPE GreenLake API and Developer Basics",
        "title": "HPE GreenLake API and Developer Basics",
        "category": "HPE GreenLake",
        "description": "Demo reference for locating API-oriented information and understanding the HPE GreenLake developer experience.",
        "keywords": "GreenLake API REST OpenAPI OAuth developer portal automation",
        "source_url": "https://developer.hpe.com/platform/hpe-greenlake/home/",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "The HPE GreenLake developer portal provides public technical resources and API documentation. HPE describes REST and OpenAPI-based interfaces and authentication mechanisms for supported services. When investigating an API issue, capture the service, endpoint, authentication context and response information."
    },
    {
        "filename": "HPE Aruba Central Monitoring and Alerts",
        "title": "HPE Aruba Central Monitoring and Alerts",
        "category": "Networking",
        "description": "Demo reference for Central monitoring, health indicators, reporting and troubleshooting workflows.",
        "keywords": "Aruba Central monitoring alerts health dashboard reporting troubleshooting",
        "source_url": "https://arubanetworking.hpe.com/techdocs/central/2.5.7/content/faqs/overview.htm",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE Aruba Central provides cloud-managed networking capabilities with monitoring, reporting and troubleshooting functions. For an alert investigation, capture the site, device, affected service, alert time and health metric before selecting a remediation path."
    },
    {
        "filename": "HPE Aruba Central Wired and Wireless Management",
        "title": "HPE Aruba Central Wired and Wireless Management",
        "category": "Networking",
        "description": "Demo reference covering the broad wired, wireless, WAN and VPN management scope of Aruba Central.",
        "keywords": "Aruba Central wired wireless WAN VPN LAN management",
        "source_url": "https://www.hpe.com/us/en/aruba-central.html.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Public HPE Aruba Central information describes unified management across supported networking domains, including wired and wireless LAN environments and WAN or VPN services. Support cases should identify the network domain and device type before applying a product-specific workflow."
    },
    {
        "filename": "HPE Aruba CX Switching Fundamentals",
        "title": "HPE Aruba CX Switching Fundamentals",
        "category": "Networking",
        "description": "Demo reference for AOS-CX switch identification, cloud management context and secure network access concepts.",
        "keywords": "Aruba CX AOS-CX switch switching network access cloud management",
        "source_url": "https://arubanetworking.hpe.com/techdocs/central/2.5.8/content/nms/landing-pages/switches.htm",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE Aruba public documentation describes AOS-CX switches as part of the networking portfolio and documents cloud management workflows through Aruba Central. Before troubleshooting a switch, verify the exact platform, operating system, software version and site role."
    },
    {
        "filename": "HPE Aruba Wi-Fi and Access Point Overview",
        "title": "HPE Aruba Wi-Fi and Access Point Overview",
        "category": "Product",
        "description": "Demo product reference for HPE Aruba Networking Wi-Fi and access point solutions.",
        "keywords": "Aruba Wi-Fi access point wireless AP networking",
        "source_url": "https://www.hpe.com/us/en/networking.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE Aruba Networking includes Wi-Fi and access point solutions within its networking portfolio. For a wireless support case, identify the access point model, management platform, site and client symptom before selecting a troubleshooting procedure."
    },
    {
        "filename": "HPE Aruba EdgeConnect SD-WAN Overview",
        "title": "HPE Aruba EdgeConnect SD-WAN Overview",
        "category": "Networking",
        "description": "Demo reference for locating EdgeConnect SD-WAN within the HPE Aruba Networking portfolio.",
        "keywords": "EdgeConnect SD-WAN Aruba WAN networking branch connectivity",
        "source_url": "https://www.hpe.com/us/en/networking.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE's networking portfolio includes EdgeConnect SD-WAN. For a WAN support case, capture the affected site, tunnel or path, connectivity symptom, device or appliance identity and time of occurrence before selecting a product-specific workflow."
    },
    {
        "filename": "HPE Aruba Networking Services Quick Guide",
        "title": "HPE Aruba Networking Services Quick Guide",
        "category": "Support",
        "description": "Demo guide showing how networking services can align to planning, deployment, operations, support and optimization.",
        "keywords": "Aruba networking services planning deployment support education optimization",
        "source_url": "https://www.hpe.com/sg/en/products/networking/hpe-aruba-networking-services.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE Aruba Networking services cover lifecycle activities such as planning and design, deployment and migration, operations and optimization. Public HPE information also describes support, education and professional services. Classify the customer request by lifecycle stage before routing it."
    },
    {
        "filename": "HPE Networking Support Case Intake Checklist",
        "title": "HPE Networking Support Case Intake Checklist",
        "category": "Support",
        "description": "Mock support intake reference for collecting product, site, symptom, time and environment details.",
        "keywords": "networking support case intake troubleshooting site device symptom severity",
        "source_url": "https://www.hpe.com/us/en/networking.html",
        "source_type": "Demo Workflow Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Demo intake guidance: collect the customer environment, exact product or device, site, observed symptom, start time, recent changes and business impact. Use the collected facts to select the correct HPE Aruba or networking product reference. This record is a mock workflow for the knowledge-base interface."
    },
    {
        "filename": "HPE GreenLake Backup and Recovery Basics",
        "title": "HPE GreenLake Backup and Recovery Basics",
        "category": "Data Protection",
        "description": "Demo reference for backup and restore use cases across supported virtual, database and storage resources.",
        "keywords": "GreenLake backup recovery VMware SQL Server array volumes restore",
        "source_url": "https://support.hpe.com/hpesc/public/docDisplay?docId=sd00003454en_us&docLocale=en_US&page=GUID-76A1AF84-7118-45D9-B2E3-F0166D955259.html",
        "source_type": "Public HPE Reference",
        "page_count": 1,
        "featured": 0,
        "text": "HPE GreenLake for Backup and Recovery is a data protection service described in public HPE documentation. Supported scenarios include backup and restore for selected VMware resources, Microsoft SQL Server resources and HPE array volumes. Verify subscription, account and required roles when investigating access or protection issues."
    },
    {
        "filename": "HPE Backup and Recovery Troubleshooting Intake",
        "title": "HPE Backup and Recovery Troubleshooting Intake",
        "category": "Data Protection",
        "description": "Mock intake checklist for backup failures, restore requests and data-protection access issues.",
        "keywords": "backup failure restore troubleshooting protection job subscription role",
        "source_url": "https://support.hpe.com/hpesc/public/docDisplay?docId=sd00003454en_us&docLocale=en_US&page=GUID-76A1AF84-7118-45D9-B2E3-F0166D955259.html",
        "source_type": "Demo Workflow Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Demo intake guidance: identify the protected resource, backup or restore job, failure time, error text, subscription and user role. For recovery incidents, confirm whether the issue affects backup creation, retention, cataloging or restore execution before routing the case."
    },
    {
        "filename": "HPE Knowledge Base Search Guide",
        "title": "HPE Knowledge Base Search Guide",
        "category": "General",
        "description": "Mock orientation guide explaining how to search the knowledge base by product, symptom, technology or support topic.",
        "keywords": "knowledge base search product symptom keyword troubleshooting guide",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Content",
        "page_count": 1,
        "featured": 0,
        "text": "Demo search guidance: use product names, platform names, management technologies, error terms and symptoms as search phrases. Examples include ProLiant iLO, Aruba Central, AOS-CX, GreenLake access, backup restore and server health. Search results combine seeded public reference records with any PDFs added later."
    },
    {
        "filename": "HPE Product Identification Quick Guide",
        "title": "HPE Product Identification Quick Guide",
        "category": "General",
        "description": "Mock guide for identifying the product family before selecting a detailed technical article.",
        "keywords": "product identification ProLiant Aruba GreenLake networking server storage",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Content",
        "page_count": 1,
        "featured": 0,
        "text": "Demo guidance: identify whether a request concerns compute, management, HPE GreenLake, networking, support or data protection. Then capture the specific product, model or service. Correct product identification reduces the chance of applying an unrelated procedure."
    },
    {
        "filename": "HPE Support Case Documentation Checklist",
        "title": "HPE Support Case Documentation Checklist",
        "category": "Procedures",
        "description": "Mock procedure for documenting a technical case consistently before escalation or handoff.",
        "keywords": "support case documentation escalation handoff symptoms timeline evidence",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Workflow Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Demo procedure: document the customer impact, affected product, exact symptom, timeline, recent changes, troubleshooting already completed and relevant evidence. For escalation, preserve the original error text and clearly state the requested next action. This is mock workflow content for demonstration."
    },
    {
        "filename": "HPE Troubleshooting Information Quality Guide",
        "title": "HPE Troubleshooting Information Quality Guide",
        "category": "Procedures",
        "description": "Mock procedure for improving search and troubleshooting quality by using precise technical terms and verified identifiers.",
        "keywords": "troubleshooting information quality serial model version error logs search terms",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Workflow Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Demo procedure: use exact product names, model identifiers, software versions and error messages when searching. Avoid relying on broad terms alone. When a result appears relevant, verify that the platform and environment match the customer case before applying its guidance."
    },
    {
        "filename": "HPE Learning and Enablement Reference",
        "title": "HPE Learning and Enablement Reference",
        "category": "Training",
        "description": "Mock training reference that groups knowledge-base learning around compute, networking, GreenLake and support fundamentals.",
        "keywords": "HPE training learning compute networking GreenLake support enablement",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Content",
        "page_count": 1,
        "featured": 0,
        "text": "Demo learning path: start with product identification, then learn the relevant management platform, common support terminology and basic troubleshooting intake. Suggested topic groups include ProLiant and iLO, Aruba Networking and Central, HPE GreenLake, and backup and recovery."
    },
    {
        "filename": "HPE Networking Learning Topics",
        "title": "HPE Networking Learning Topics",
        "category": "Training",
        "description": "Mock learning index for Aruba Central, switches, Wi-Fi, SD-WAN and networking services.",
        "keywords": "Aruba Central switches Wi-Fi SD-WAN training learning networking",
        "source_url": "https://www.hpe.com/us/en/networking.html",
        "source_type": "Demo Content",
        "page_count": 1,
        "featured": 0,
        "text": "Demo training index: networking learners can organize study around Aruba Central, wired switching, Wi-Fi access points, SD-WAN and networking lifecycle services. The correct technical article should always be selected using the specific product and version involved."
    },
    {
        "filename": "HPE Reference Content Disclaimer",
        "title": "HPE Reference Content Disclaimer",
        "category": "Policies",
        "description": "Demo policy note explaining that seeded records are reference summaries and not proprietary internal procedures.",
        "keywords": "public reference demo content disclaimer proprietary internal policy",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Content",
        "page_count": 1,
        "featured": 0,
        "text": "This knowledge-base demo uses concise summaries of publicly available HPE information plus clearly labeled mock workflow records. It is not a substitute for customer-specific contracts, support entitlements, internal procedures or official product documentation. Verify current official documentation before performing production changes."
    },
    {
        "filename": "HPE Knowledge Base Content Standards",
        "title": "HPE Knowledge Base Content Standards",
        "category": "Policies",
        "description": "Mock content standard for keeping knowledge-base records concise, searchable, traceable and easy to maintain.",
        "keywords": "knowledge base content standard title keywords source category maintenance",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Content",
        "page_count": 1,
        "featured": 0,
        "text": "Demo content standard: every record should have a clear title, category, concise description, searchable keywords and a source or content-type label. Public reference summaries should point users toward the original source. Mock workflows should be visibly identified as demo content."
    },
    {
        "filename": "HPE Customer Impact and Severity Basics",
        "title": "HPE Customer Impact and Severity Basics",
        "category": "Support",
        "description": "Mock support reference for capturing business impact and urgency without replacing contractual severity definitions.",
        "keywords": "support severity customer impact urgency escalation business impact",
        "source_url": "https://www.hpe.com/",
        "source_type": "Demo Workflow Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Demo guidance: record the affected service, number of impacted users or systems, business impact, start time and current workaround. Do not infer a contractual support severity from this mock record; use the customer's applicable support agreement and official case process."
    },
    {
        "filename": "HPE Compute Troubleshooting Intake",
        "title": "HPE Compute Troubleshooting Intake",
        "category": "Procedures",
        "description": "Mock compute intake reference for server model, generation, iLO state, symptoms and recent changes.",
        "keywords": "ProLiant troubleshooting intake model generation iLO server health error",
        "source_url": "https://www.hpe.com/us/en/products/compute/proliant.html",
        "source_type": "Demo Workflow Reference",
        "page_count": 1,
        "featured": 0,
        "text": "Demo procedure: for a compute issue, capture the exact ProLiant model, generation, iLO version where applicable, observed health state, error text, recent hardware or software changes and customer impact. Use official product documentation for model-specific remediation."
    },

]


def seed_mock_documents():
    conn = db()
    now = datetime.now().isoformat(timespec="seconds")

    # Seed each public HPE reference independently. This means the demo
    # data will still be added if an existing knowledge_base.db already
    # contains older uploaded documents.
    for item in MOCK_DOCS:
        file_hash = hashlib.sha256(
            f"PUBLIC-HPE-MOCK::{item['filename']}".encode()
        ).hexdigest()

        existing = conn.execute(
            "SELECT id FROM documents WHERE file_hash = ?",
            (file_hash,),
        ).fetchone()

        if existing:
            continue

        cursor = conn.execute(
            """
            INSERT INTO documents
            (filename, stored_path, file_hash, category, title, description,
             keywords, source_url, source_type, page_count, file_size,
             uploaded_at, indexed_at, status, featured)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                item["filename"],
                "",
                file_hash,
                item["category"],
                item["title"],
                item["description"],
                item["keywords"],
                item["source_url"],
                item["source_type"],
                item["page_count"],
                0,
                now,
                now,
                "Indexed",
                item["featured"],
            ),
        )

        document_id = cursor.lastrowid

        chunks = split_text(
            item["text"],
            chunk_size=160,
            overlap=25,
        )

        for chunk_index, chunk in enumerate(chunks):
            conn.execute(
                """
                INSERT INTO chunks
                (document_id, page_number, chunk_index, text)
                VALUES (?, ?, ?, ?)
                """,
                (
                    document_id,
                    1,
                    chunk_index,
                    chunk,
                ),
            )

    conn.commit()
    conn.close()


def split_text(text, chunk_size=1100, overlap=180):
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


seed_mock_documents()


# ============================================================
# DATABASE HELPERS
# ============================================================

@st.cache_data(ttl=60, show_spinner=False)
def get_documents():
    conn = db()
    rows = conn.execute(
        "SELECT * FROM documents ORDER BY uploaded_at DESC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


@st.cache_data(ttl=60, show_spinner=False)
def get_categories():
    conn = db()
    rows = conn.execute(
        "SELECT DISTINCT category FROM documents ORDER BY category"
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
            d.title,
            d.description,
            d.keywords,
            d.source_url,
            d.source_type,
            d.stored_path,
            d.page_count
        FROM chunks c
        JOIN documents d ON d.id = c.document_id
        ORDER BY c.id
        """
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


@st.cache_data(ttl=60, show_spinner=False)
def get_document_by_id(document_id):
    conn = db()
    row = conn.execute(
        "SELECT * FROM documents WHERE id = ?",
        (document_id,),
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def clear_caches():
    for fn in (
        get_documents,
        get_categories,
        get_document_by_id,
        build_search_index,
        search_documents,
        find_document_matches,
        render_pdf_page,
        render_pdf_page_highlighted,
    ):
        try:
            fn.clear()
        except Exception:
            pass


# ============================================================
# SEARCH ENGINE
# ============================================================

@st.cache_data(ttl=60, show_spinner=False)
def build_search_index():
    rows = get_all_chunks()

    if not rows:
        return None, None, []

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

    return vectorizer, matrix, rows


def make_snippet(text, query, radius=250):
    text_clean = re.sub(r"\s+", " ", text).strip()

    if not query:
        return text_clean[:radius]

    terms = [
        t.lower()
        for t in re.findall(r"\w+", query)
        if len(t) > 2
    ]

    lower = text_clean.lower()
    positions = []

    for term in terms:
        pos = lower.find(term)
        if pos >= 0:
            positions.append(pos)

    if not positions:
        return text_clean[:radius] + ("..." if len(text_clean) > radius else "")

    center = min(positions)
    start = max(0, center - radius // 2)
    end = min(len(text_clean), start + radius)

    return (
        ("..." if start > 0 else "")
        + text_clean[start:end]
        + ("..." if end < len(text_clean) else "")
    )


def exact_passage(text, query, max_chars=1100):
    normalized = re.sub(r"\s+", " ", text).strip()

    if not normalized:
        return ""

    sentences = re.split(r"(?<=[.!?])\s+", normalized)

    terms = [
        t.lower()
        for t in re.findall(r"[A-Za-z0-9]+", query)
        if len(t) > 2
    ]

    scored = []

    for index, sentence in enumerate(sentences):
        low = sentence.lower()
        score = sum(low.count(term) for term in terms)
        if score:
            scored.append((score, index))

    if not scored:
        return normalized[:max_chars]

    scored.sort(key=lambda x: (-x[0], x[1]))

    selected = set()

    for _, index in scored[:4]:
        selected.add(index)
        if index + 1 < len(sentences):
            selected.add(index + 1)

    passage = " ".join(
        sentences[index]
        for index in sorted(selected)
    )

    if len(passage) > max_chars:
        passage = passage[:max_chars].rsplit(" ", 1)[0] + "..."

    return passage


@st.cache_data(ttl=30, show_spinner=False)
def search_documents(query, category="All Categories", top_k=10):
    query = query.strip()

    if not query:
        return []

    vectorizer, matrix, rows = build_search_index()

    if vectorizer is None:
        return []

    query_vector = vectorizer.transform([query])
    scores = cosine_similarity(query_vector, matrix).flatten()

    best_by_document = {}

    for index, score in enumerate(scores):
        row = rows[index]

        if category != "All Categories" and row["category"] != category:
            continue

        if score <= 0:
            continue

        result = {
            **row,
            "score": float(score),
            "snippet": make_snippet(row["text"], query),
            "exact_passage": exact_passage(row["text"], query),
        }

        current = best_by_document.get(row["document_id"])

        if current is None or result["score"] > current["score"]:
            best_by_document[row["document_id"]] = result

    results = list(best_by_document.values())
    results.sort(key=lambda x: x["score"], reverse=True)

    return results[:top_k]


# ============================================================
# PDF FUNCTIONS
# ============================================================

def clean_text(text):
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def make_hash(data):
    return hashlib.sha256(data).hexdigest()


def extract_pdf(pdf_bytes):
    pages = []

    with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
        for page_number, page in enumerate(pdf, start=1):
            text = clean_text(page.get_text("text"))
            pages.append((page_number, text))

        return pages, len(pdf)


def save_pdf(file_name, pdf_bytes, file_hash):
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", file_name)
    destination = PDF_DIR / f"{file_hash[:12]}_{safe_name}"
    destination.write_bytes(pdf_bytes)
    return str(destination)


def add_document(file_name, pdf_bytes, category):
    file_hash = make_hash(pdf_bytes)

    conn = db()

    existing = conn.execute(
        "SELECT id FROM documents WHERE file_hash = ?",
        (file_hash,),
    ).fetchone()

    if existing:
        conn.close()
        return False, "This PDF has already been uploaded."

    try:
        pages, page_count = extract_pdf(pdf_bytes)
    except Exception as e:
        conn.close()
        return False, f"Could not read PDF: {e}"

    stored_path = save_pdf(file_name, pdf_bytes, file_hash)
    now = datetime.now().isoformat(timespec="seconds")

    cursor = conn.execute(
        """
        INSERT INTO documents
        (filename, stored_path, file_hash, category, title, description,
         keywords, source_url, source_type, page_count, file_size,
         uploaded_at, indexed_at, status, featured)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            file_name,
            stored_path,
            file_hash,
            category,
            Path(file_name).stem,
            f"Uploaded PDF: {file_name}",
            category,
            "",
            "Uploaded PDF",
            page_count,
            len(pdf_bytes),
            now,
            now,
            "Indexed",
            0,
        ),
    )

    document_id = cursor.lastrowid

    for page_number, page_text in pages:
        for chunk_index, chunk in enumerate(split_text(page_text)):
            conn.execute(
                """
                INSERT INTO chunks
                (document_id, page_number, chunk_index, text)
                VALUES (?, ?, ?, ?)
                """,
                (document_id, page_number, chunk_index, chunk),
            )

    conn.commit()
    conn.close()

    return True, f"{file_name} indexed successfully."


def delete_document(document_id):
    conn = db()

    row = conn.execute(
        "SELECT stored_path FROM documents WHERE id = ?",
        (document_id,),
    ).fetchone()

    if row and row["stored_path"]:
        try:
            Path(row["stored_path"]).unlink(missing_ok=True)
        except Exception:
            pass

    conn.execute(
        "DELETE FROM chunks WHERE document_id = ?",
        (document_id,),
    )

    conn.execute(
        "DELETE FROM documents WHERE id = ?",
        (document_id,),
    )

    conn.commit()
    conn.close()


def format_bytes(value):
    value = float(value or 0)

    if value < 1024:
        return f"{value:.0f} B"
    if value < 1024**2:
        return f"{value / 1024:.1f} KB"
    if value < 1024**3:
        return f"{value / 1024**2:.1f} MB"

    return f"{value / 1024**3:.1f} GB"


@st.cache_data(ttl=600, show_spinner=False)
def render_pdf_page(document_path, page_number, scale=1.6):
    try:
        with fitz.open(document_path) as pdf:
            if page_number < 1 or page_number > len(pdf):
                return None

            page = pdf[page_number - 1]

            pix = page.get_pixmap(
                matrix=fitz.Matrix(float(scale), float(scale)),
                alpha=False,
            )

            return pix.tobytes("png")

    except Exception:
        return None


@st.cache_data(ttl=600, show_spinner=False)
def render_pdf_page_highlighted(
    document_path,
    page_number,
    query="",
    scale=1.6,
):
    try:
        with fitz.open(document_path) as pdf:
            if page_number < 1 or page_number > len(pdf):
                return None

            page = pdf[page_number - 1]

            terms = [
                t
                for t in re.findall(r"[A-Za-z0-9]+", query)
                if len(t) > 2
            ]

            for term in terms[:12]:
                try:
                    for rect in page.search_for(term):
                        annot = page.add_highlight_annot(rect)
                        annot.update()
                except Exception:
                    continue

            pix = page.get_pixmap(
                matrix=fitz.Matrix(float(scale), float(scale)),
                alpha=False,
            )

            return pix.tobytes("png")

    except Exception:
        return None


@st.cache_data(ttl=120, show_spinner=False)
def find_document_matches(document_path, query, limit=8):
    try:
        terms = [
            t.lower()
            for t in re.findall(r"[A-Za-z0-9]+", query)
            if len(t) > 2
        ]

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

                if score:
                    matches.append(
                        {
                            "page": page_number,
                            "score": score,
                            "snippet": make_snippet(raw, query, 190),
                        }
                    )

        matches.sort(
            key=lambda x: (-x["score"], x["page"])
        )

        return matches[:limit]

    except Exception:
        return []


# ============================================================
# SIGNED BROWSER ACCESS
# ============================================================

def get_serializer():
    if URLSafeTimedSerializer is None:
        return None

    return URLSafeTimedSerializer(
        TOKEN_SECRET,
        salt="hpe-knowledge-base-access",
    )


def create_browser_token():
    serializer = get_serializer()

    if serializer is None:
        return ""

    return serializer.dumps(
        {
            "authorized": True,
        }
    )


def validate_browser_token(token):
    if not token:
        return False

    serializer = get_serializer()

    if serializer is None:
        return False

    try:
        payload = serializer.loads(token)

        return bool(
            payload.get("authorized")
        )

    except Exception:
        return False


def browser_is_authorized():
    if st.session_state.access_authorized:
        return True

    try:
        token = st.query_params.get("kb_access", "")
    except Exception:
        token = ""

    if validate_browser_token(token):
        st.session_state.access_authorized = True
        return True

    return False


def authorize_browser():
    token = create_browser_token()

    if not token:
        return False

    st.query_params["kb_access"] = token
    st.session_state.access_authorized = True

    return True


def clear_browser_access():
    st.session_state.access_authorized = False

    try:
        st.query_params.clear()
    except Exception:
        pass


# ============================================================
# ACCESS GATE
# ============================================================

def render_access_gate():
    st.markdown(
        """
        <div class="access-wrap">
          <div class="access-logo">HP<span>E</span></div>
          <div class="access-title">Knowledge Base Access</div>
          <div class="access-sub">
            Enter the access code to continue.
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not ACCESS_CODE:
        st.error(
            "ACCESS_CODE is not configured. Add it to Streamlit Secrets."
        )
        st.stop()

    _, center, _ = st.columns([1, 2, 1])

    with center:
        with st.form("access_form"):
            code = st.text_input(
                "Access Code",
                type="password",
                placeholder="Enter access code",
                label_visibility="collapsed",
            )

            submit = st.form_submit_button(
                "Access Knowledge Base",
                type="primary",
                use_container_width=True,
            )

        if submit:
            if hmac.compare_digest(
                code.strip(),
                ACCESS_CODE,
            ):
                if authorize_browser():
                    st.rerun()
                else:
                    st.error(
                        "Could not create browser authorization."
                    )
            else:
                st.error("Invalid access code.")

        st.caption(
            "The access code is not stored in the URL. "
            "A signed browser authorization token is used instead."
        )


if not browser_is_authorized():
    render_access_gate()
    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

navigation = [
    ("⌂", "Home"),
    ("⌕", "Search"),
    ("▣", "Browse All"),
    ("▦", "Categories"),
    ("★", "Featured"),
    ("◷", "Recent"),
    ("☁", "Upload PDF"),
    ("⚙", "Manage Content"),
    ("▤", "Analytics"),
    ("?", "Help"),
]

with st.sidebar:
    st.markdown(
        '<div class="sidebar-logo">HP<span class="e">E</span></div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="sidebar-space"></div>',
        unsafe_allow_html=True,
    )

    for icon, label in navigation:
        if st.button(
            f"{icon}    {label}",
            key=f"nav_{label}",
            use_container_width=True,
        ):
            if label == "Manage Content" and not st.session_state.admin_authenticated:
                st.session_state.page = "Admin Login"
            else:
                st.session_state.page = label

            st.session_state.selected_result_id = None
            st.rerun()

    st.markdown(
        """
        <div class="sidebar-footer">
          <div class="big">HPE</div>
          <div class="small">Knowledge Base</div>
          <div class="version">Public Reference Demo · v2.0</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# HEADER
# ============================================================

page_title = st.session_state.page

h1, h2, h3, h4 = st.columns(
    [1.15, 3.05, .55, 1.15],
    vertical_alignment="center",
)

with h1:
    st.markdown(
        f'<div class="page-heading">{page_title}</div>',
        unsafe_allow_html=True,
    )

with h2:
    top_search = st.text_input(
        "Header search",
        value=st.session_state.search_query,
        placeholder="Search for topics, keywords, or questions...",
        label_visibility="collapsed",
        key="top_search",
    )

with h3:
    if st.button(
        "⚙",
        key="header_admin",
        help="Admin access",
        use_container_width=True,
    ):
        st.session_state.page = (
            "Manage Content"
            if st.session_state.admin_authenticated
            else "Admin Login"
        )
        st.rerun()

with h4:
    role = (
        "Administrator"
        if st.session_state.admin_authenticated
        else "Authorized User"
    )

    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:9px;">
          <div class="avatar">👩🏻</div>
          <div>
            <div class="profile-name">Knowledge User</div>
            <div class="profile-role">{role}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


if top_search.strip():
    st.session_state.search_query = top_search
    st.session_state.page = "Search"


# ============================================================
# ADMIN LOGIN
# ============================================================

def admin_login_page():
    st.markdown(
        """
        <div class="admin-card">
          <div class="admin-icon">⚙</div>
          <div class="admin-title-big">Admin Access</div>
          <div class="admin-subtitle">
            Enter the administrator PIN to manage the Knowledge Base.
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.form("admin_login"):
        pin = st.text_input(
            "Admin PIN",
            type="password",
            placeholder="Enter admin PIN",
        )

        login = st.form_submit_button(
            "Unlock Admin",
            type="primary",
            use_container_width=True,
        )

    if login:
        if hmac.compare_digest(
            pin.strip(),
            ADMIN_PIN,
        ):
            st.session_state.admin_authenticated = True
            st.session_state.page = "Manage Content"
            st.rerun()
        else:
            st.error("Incorrect administrator PIN.")

    if st.button("← Back to Home"):
        st.session_state.page = "Home"
        st.rerun()


# ============================================================
# HOME
# ============================================================

def home_page():
    documents = get_documents()
    category_counts = {}

    for document in documents:
        category = document["category"]
        category_counts[category] = (
            category_counts.get(category, 0) + 1
        )

    st.markdown(
        """
        <div class="hero">
          <h1>Find the answers you need</h1>
          <div class="hero-sub">
            Search our knowledge base, explore topics, or browse by category.
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    hero_search = st.text_input(
        "Hero search",
        placeholder="Search for solutions, guides, or keywords...",
        label_visibility="collapsed",
        key="hero_search",
    )

    if hero_search.strip():
        st.session_state.search_query = hero_search
        st.session_state.page = "Search"
        st.rerun()

    st.markdown(
        """
        <div class="popular-row">
          <span class="popular-text">Popular searches:</span>
          <span class="popular-pill">ProLiant</span>
          <span class="popular-pill">iLO</span>
          <span class="popular-pill">GreenLake</span>
          <span class="popular-pill">Aruba Central</span>
          <span class="popular-pill">CX Switches</span>
          <span class="popular-pill">Backup & Recovery</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="metrics"></div>',
        unsafe_allow_html=True,
    )

    metric_data = [
        ("📖", len(documents), "Total Documents", "#dffaf1"),
        ("▦", len(category_counts), "Categories", "#e4f1ff"),
        (
            "★",
            sum(1 for d in documents if d["featured"]),
            "Featured",
            "#fff4d7",
        ),
        (
            "☁",
            sum(1 for d in documents if d["source_type"] == "Uploaded PDF"),
            "Uploaded PDFs",
            "#f0e9ff",
        ),
    ]

    cols = st.columns(4, gap="small")

    for col, (icon, number, label, background) in zip(
        cols,
        metric_data,
    ):
        with col:
            st.markdown(
                f"""
                <div class="metric">
                  <div class="metric-icon" style="background:{background}">
                    {icon}
                  </div>
                  <div class="metric-body">
                    <div class="metric-number">{number}</div>
                    <div class="metric-label">{label}</div>
                  </div>
                  <div class="metric-arrow">›</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.write("")

    # CATEGORY CARD
    st.markdown(
        """
        <div class="white-card">
          <div class="section-head">
            <div class="section-title">
              ▦ &nbsp;Browse by Category
            </div>
            <div class="section-link">
              Knowledge Base
            </div>
          </div>
        """,
        unsafe_allow_html=True,
    )

    icons = {
        "Compute": ("▣", "#dffaf1"),
        "Management": ("⚙", "#e4f1ff"),
        "HPE GreenLake": ("☁", "#dffaf1"),
        "Networking": ("⌁", "#e4f1ff"),
        "Support": ("♢", "#fff4d7"),
        "Data Protection": ("▤", "#f0e9ff"),
        "Product": ("▣", "#e4f1ff"),
        "General": ("▤", "#edf2f5"),
    }

    category_items = list(category_counts.items())[:8]

    for start in range(
        0,
        len(category_items),
        4,
    ):
        row = category_items[start:start + 4]
        category_cols = st.columns(4, gap="small")

        for col, (category, count) in zip(
            category_cols,
            row,
        ):
            icon, background = icons.get(
                category,
                ("▤", "#edf2f5"),
            )

            with col:
                st.markdown(
                    f"""
                    <div class="category-tile">
                      <div class="category-icon"
                           style="background:{background}">
                        {icon}
                      </div>
                      <div>
                        <div class="category-name">
                          {category}
                        </div>
                        <div class="category-count">
                          {count} documents
                        </div>
                      </div>
                      <div class="category-arrow">›</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                if st.button(
                    "Open",
                    key=f"home_category_{category}",
                    use_container_width=True,
                ):
                    st.session_state.search_query = category
                    st.session_state.page = "Search"
                    st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    st.write("")

    recent = documents[:5]

    featured = next(
        (
            d
            for d in documents
            if d["featured"]
        ),
        documents[0] if documents else None,
    )

    left, right = st.columns(
        [1.08, .92],
        gap="small",
    )

    # RECENT DOCUMENTS
    with left:
        st.markdown(
            """
            <div class="white-card">
              <div class="section-head">
                <div class="section-title">
                  ▤ &nbsp;Recent Documents
                </div>
                <div class="section-link">
                  Latest
                </div>
              </div>
            """,
            unsafe_allow_html=True,
        )

        for document in recent:
            r1, r2, r3 = st.columns(
                [.55, 5.0, 1.3],
                vertical_alignment="center",
            )

            with r1:
                st.markdown(
                    '<div class="pdf-icon">PDF</div>',
                    unsafe_allow_html=True,
                )

            with r2:
                st.markdown(
                    f"""
                    <div class="recent-title">
                      {document["title"]}
                    </div>
                    <div class="recent-meta">
                      {document["category"]}
                      &nbsp;·&nbsp;
                      {document["source_type"]}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with r3:
                st.markdown(
                    f"""
                    <div class="recent-date">
                      {document["uploaded_at"][:10]}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            if st.button(
                "Open",
                key=f"recent_{document['id']}",
                use_container_width=True,
            ):
                st.session_state.search_query = document["title"]
                st.session_state.page = "Search"
                st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

    # FEATURED DOCUMENT
    with right:
        st.markdown(
            """
            <div class="white-card">
              <div class="section-head">
                <div class="section-title">
                  ★ &nbsp;Featured Document
                </div>
              </div>
            """,
            unsafe_allow_html=True,
        )

        if featured:
            st.markdown(
                f"""
                <div class="viewer">
                  <div class="paper">
                    <div class="paper-logo">
                      HP<span>E</span>
                    </div>
                    <div class="paper-title">
                      {featured["title"]}
                    </div>
                    <div class="paper-rule"></div>
                    <div class="paper-building"></div>
                    <div class="paper-small">
                      Public HPE reference · {featured["category"]}
                    </div>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown(
                f"""
                <div style="font-size:13px;font-weight:700;margin-top:10px;">
                  {featured["title"]}
                </div>
                <div style="font-size:10px;color:#707c87;margin-top:3px;">
                  {featured["category"]}
                  &nbsp;·&nbsp;
                  Public HPE Reference
                </div>
                """,
                unsafe_allow_html=True,
            )

            if st.button(
                "Open Document →",
                key="open_featured",
                type="primary",
                use_container_width=True,
            ):
                st.session_state.search_query = featured["title"]
                st.session_state.page = "Search"
                st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# SEARCH PAGE
# ============================================================

def search_page():
    left_search, filter_col = st.columns(
        [7.4, 1.0],
        gap="small",
    )

    with left_search:
        with st.form(
            "search_form",
            clear_on_submit=False,
        ):
            input_col, button_col = st.columns(
                [6.4, 1.0],
                gap="small",
            )

            with input_col:
                query = st.text_input(
                    "Search",
                    value=st.session_state.search_query,
                    placeholder=(
                        "What should I check or find "
                        "in the knowledge base?"
                    ),
                    label_visibility="collapsed",
                )

            with button_col:
                search_clicked = st.form_submit_button(
                    "Search",
                    type="primary",
                    use_container_width=True,
                )

    with filter_col:
        with st.popover(
            "☷ Filters",
            use_container_width=True,
        ):
            category = st.selectbox(
                "Category",
                ["All Categories"] + get_categories(),
            )

            top_k = st.selectbox(
                "Results",
                [5, 10, 20],
                index=1,
            )

    if search_clicked:
        st.session_state.search_query = query
        st.session_state.selected_result_id = None
        st.session_state.viewer_page = 1

    active_query = st.session_state.search_query.strip()

    if not active_query:
        documents = get_documents()
        page_count = sum(
            int(d["page_count"] or 0)
            for d in documents
        )

        st.markdown(
            f"""
            <div class="welcome-card">
              <div class="welcome-icon">⌕</div>
              <div class="welcome-title">
                Search your knowledge base
              </div>
              <div class="welcome-text">
                Search public HPE reference data and your uploaded PDFs
                for products, procedures, troubleshooting terms,
                management technologies and support topics.
              </div>
              <div class="welcome-stats">
                <span>
                  <b>{len(documents)}</b> documents
                </span>
                <span>
                  <b>{page_count:,}</b> indexed pages
                </span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        return

    signature = (
        active_query,
        category,
        top_k,
    )

    if (
        st.session_state.search_signature
        != signature
    ):
        st.session_state.search_results = (
            search_documents(
                active_query,
                category=category,
                top_k=top_k,
            )
        )

        st.session_state.search_signature = signature
        st.session_state.selected_result_id = None

    results = st.session_state.search_results

    if not results:
        st.warning(
            "No matching knowledge-base content found. "
            "Try a product name, feature, or troubleshooting keyword."
        )
        return

    best = results[0]

    selected_id = (
        st.session_state.selected_result_id
    )

    selected = next(
        (
            result
            for result in results
            if result["id"] == selected_id
        ),
        best,
    )

    if selected_id != selected["id"]:
        st.session_state.selected_result_id = selected["id"]
        st.session_state.viewer_page = int(
            selected["page_number"]
        )

    selected_document = get_document_by_id(
        selected["document_id"]
    )

    total_pages = int(
        selected_document["page_count"]
        if selected_document
        else 1
    )

    if (
        st.session_state.viewer_page < 1
        or st.session_state.viewer_page > total_pages
    ):
        st.session_state.viewer_page = int(
            selected["page_number"]
        )

    viewer_page = st.session_state.viewer_page

    st.markdown(
        f"""
        <div class="search-count">
          {len(results)} result(s)
          · Highest match first
          · Source text remains tied to the selected document
        </div>
        """,
        unsafe_allow_html=True,
    )

    result_col, source_col = st.columns(
        [.82, 1.8],
        gap="large",
    )

    # ---------------- RESULTS ----------------

    with result_col:
        st.markdown(
            '<div class="panel-title">Search Results</div>',
            unsafe_allow_html=True,
        )

        for index, result in enumerate(results):
            best_result = index == 0
            score = min(
                99,
                max(
                    1,
                    round(
                        result["score"] * 100
                    ),
                ),
            )

            with st.container(
                key=f"search_result_{index}",
                border=True,
            ):
                st.markdown(
                    f"""
                    <div class="result-label">
                      {"Highest Match" if best_result else "Search Result"}
                      <span class="result-score-pill">
                        {score}%
                      </span>
                    </div>

                    <div class="result-filename">
                      {result["filename"]}
                    </div>

                    <div class="result-page">
                      Page {result["page_number"]}
                      · {result["category"]}
                    </div>

                    <div class="result-snippet">
                      {result["snippet"]}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                if st.button(
                    "Open Result",
                    key=f"open_result_{result['id']}",
                    use_container_width=True,
                ):
                    st.session_state.selected_result_id = (
                        result["id"]
                    )

                    st.session_state.viewer_page = (
                        int(result["page_number"])
                    )

                    st.rerun()

        # ---------------- MATCHES IN SELECTED PDF ----------------

        st.markdown(
            """
            <div class="match-panel">
              <div class="match-panel-title">
                Matches in this PDF
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if selected["stored_path"]:
            matches = find_document_matches(
                selected["stored_path"],
                active_query,
            )

            if matches:
                for match in matches:
                    if st.button(
                        f"Page {match['page']}",
                        key=(
                            f"match_{selected['id']}_"
                            f"{match['page']}"
                        ),
                        use_container_width=True,
                    ):
                        st.session_state.viewer_page = (
                            match["page"]
                        )
                        st.rerun()

                    st.markdown(
                        f"""
                        <div class="match-item">
                          <div class="match-item-page">
                            Page {match["page"]}
                            · {match["score"]} match(es)
                          </div>
                          <div class="match-item-text">
                            {match["snippet"]}
                          </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
            else:
                st.caption(
                    "This seeded public reference is a searchable "
                    "knowledge record rather than a local PDF."
                )

    # ---------------- SOURCE / VIEWER ----------------

    with source_col:
        st.markdown(
            f"""
            <div class="reader-toolbar">
              <div class="reader-title">
                ▣ {selected["filename"]}
              </div>

              <div class="reader-meta">
                {selected["category"]}
                · {selected["source_type"]}
                · Page {viewer_page} of {total_pages}
              </div>

              <div class="reader-match">
                Search: “{active_query}”
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # PUBLIC MOCK RECORD VIEW
        if not selected["stored_path"]:
            st.markdown(
                f"""
                <div class="white-card"
                     style="min-height:420px;padding:24px;">

                  <div style="
                      font-size:13px;
                      color:#087c63;
                      font-weight:700;
                      margin-bottom:10px;">
                    PUBLIC HPE REFERENCE
                  </div>

                  <div style="
                      font-size:23px;
                      font-weight:750;
                      color:#102238;">
                    {selected["title"]}
                  </div>

                  <div style="
                      margin-top:9px;
                      color:#687b87;
                      font-size:12px;">
                    Category: {selected["category"]}
                  </div>

                  <div style="
                      margin-top:18px;
                      padding:16px;
                      border-left:4px solid #01a982;
                      background:#f2fbf8;
                      color:#233e4e;
                      line-height:1.6;
                      font-size:14px;">
                    {selected["exact_passage"]}
                  </div>

                  <div style="
                      margin-top:20px;
                      color:#102238;
                      font-size:13px;
                      font-weight:700;">
                    Description
                  </div>

                  <div style="
                      margin-top:6px;
                      color:#526b79;
                      font-size:13px;
                      line-height:1.55;">
                    {selected["description"]}
                  </div>

                </div>
                """,
                unsafe_allow_html=True,
            )

            if selected["source_url"]:
                st.link_button(
                    "Open official HPE source ↗",
                    selected["source_url"],
                    use_container_width=False,
                )

        # ACTUAL PDF VIEWER
        else:
            zoom = st.selectbox(
                "Zoom",
                [100, 125, 150, 175, 200],
                index=2,
                format_func=lambda x: f"{x}%",
                key="pdf_zoom",
                label_visibility="collapsed",
            )

            render_scale = (
                zoom / 100 * 1.05
            )

            image = (
                render_pdf_page_highlighted(
                    selected["stored_path"],
                    viewer_page,
                    active_query,
                    scale=render_scale,
                )
            )

            if image:
                display_width = {
                    100: 700,
                    125: 820,
                    150: 930,
                    175: 1040,
                    200: 1180,
                }.get(
                    int(zoom),
                    930,
                )

                with st.container(
                    key="pdf_viewer_shell"
                ):
                    with st.container(
                        key="pdf_prev_wrap"
                    ):
                        if st.button(
                            "‹",
                            disabled=(
                                viewer_page <= 1
                            ),
                            key="pdf_prev",
                            help="Previous page",
                        ):
                            st.session_state.viewer_page = (
                                viewer_page - 1
                            )
                            st.rerun()

                    with st.container(
                        key="pdf_next_wrap"
                    ):
                        if st.button(
                            "›",
                            disabled=(
                                viewer_page >= total_pages
                            ),
                            key="pdf_next",
                            help="Next page",
                        ):
                            st.session_state.viewer_page = (
                                viewer_page + 1
                            )
                            st.rerun()

                    st.image(
                        image,
                        width=display_width,
                    )

            else:
                st.error(
                    "Unable to render this PDF page."
                )

        # EXACT SEARCH PASSAGE
        st.markdown(
            f"""
            <div class="match-panel">
              <div class="match-panel-title">
                Relevant information
              </div>

              <div style="
                  color:#385362;
                  font-size:12px;
                  line-height:1.55;">
                {selected["exact_passage"]}
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# BROWSE ALL
# ============================================================

def browse_all_page():
    documents = get_documents()

    st.markdown(
        """
        <div class="section-title">
          Browse All Knowledge
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption(
        "Public HPE references and uploaded PDF records."
    )

    for document in documents:
        with st.container(border=True):
            left, right = st.columns(
                [5.5, 1],
                vertical_alignment="center",
            )

            with left:
                st.markdown(
                    f"**{document['title']}**"
                )

                st.caption(
                    f"{document['category']} · "
                    f"{document['source_type']} · "
                    f"{document['uploaded_at'][:10]}"
                )

                st.write(
                    document["description"]
                )

            with right:
                if st.button(
                    "Open",
                    key=f"browse_{document['id']}",
                    use_container_width=True,
                ):
                    st.session_state.search_query = (
                        document["title"]
                    )
                    st.session_state.page = "Search"
                    st.rerun()


# ============================================================
# CATEGORIES
# ============================================================

def categories_page():
    counts = {}

    for document in get_documents():
        category = document["category"]
        counts[category] = (
            counts.get(category, 0) + 1
        )

    st.markdown(
        '<div class="section-title">Browse by Category</div>',
        unsafe_allow_html=True,
    )

    st.write("")

    for category, count in counts.items():
        if st.button(
            f"{category}   ·   {count} documents",
            key=f"cat_page_{category}",
            use_container_width=True,
        ):
            st.session_state.search_query = category
            st.session_state.page = "Search"
            st.rerun()


# ============================================================
# FEATURED
# ============================================================

def featured_page():
    featured = [
        d
        for d in get_documents()
        if d["featured"]
    ]

    st.markdown(
        '<div class="section-title">Featured HPE References</div>',
        unsafe_allow_html=True,
    )

    st.write("")

    for document in featured:
        with st.container(border=True):
            st.markdown(
                f"**{document['title']}**"
            )

            st.caption(
                f"{document['category']} · "
                f"{document['source_type']}"
            )

            st.write(
                document["description"]
            )

            if st.button(
                "Open",
                key=f"featured_{document['id']}",
            ):
                st.session_state.search_query = (
                    document["title"]
                )
                st.session_state.page = "Search"
                st.rerun()


# ============================================================
# RECENT
# ============================================================

def recent_page():
    st.markdown(
        '<div class="section-title">Recent Documents</div>',
        unsafe_allow_html=True,
    )

    st.write("")

    for document in get_documents()[:15]:
        st.markdown(
            f"""
            <div class="recent-row">
              <div class="pdf-icon">PDF</div>
              <div>
                <div class="recent-title">
                  {document["title"]}
                </div>
                <div class="recent-meta">
                  {document["category"]}
                  · {document["source_type"]}
                </div>
              </div>
              <div class="recent-date">
                {document["uploaded_at"][:10]}
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# UPLOAD PDF
# ============================================================

def upload_page():
    st.markdown(
        """
        <div class="admin-bar">
          <div class="admin-title">
            Add PDF to Knowledge Base
          </div>
          <div class="admin-sub">
            Upload a PDF, extract its text, index it and make it searchable.
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    category = st.selectbox(
        "Category",
        [
            "General",
            "Compute",
            "Management",
            "HPE GreenLake",
            "Networking",
            "Support",
            "Data Protection",
            "Product",
            "Policies",
            "Procedures",
            "Training",
            "Other",
        ],
    )

    uploaded = st.file_uploader(
        "Upload PDF",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if uploaded:
        st.caption(
            f"{len(uploaded)} PDF file(s) selected."
        )

    if st.button(
        "Upload & Index",
        type="primary",
        use_container_width=True,
    ):
        if not uploaded:
            st.warning(
                "Select at least one PDF."
            )
            return

        progress = st.progress(0)
        success = 0

        for index, file in enumerate(uploaded):
            ok, message = add_document(
                file.name,
                file.getvalue(),
                category,
            )

            if ok:
                success += 1
                st.success(message)
            else:
                st.warning(message)

            progress.progress(
                (index + 1) / len(uploaded)
            )

        clear_caches()

        st.success(
            f"Completed. {success} document(s) indexed."
        )

        st.session_state.search_query = ""
        st.session_state.search_results = []
        st.session_state.search_signature = None


# ============================================================
# ADMIN / MANAGE CONTENT
# ============================================================

def manage_page():
    if not st.session_state.admin_authenticated:
        st.session_state.page = "Admin Login"
        st.rerun()

    st.markdown(
        """
        <div class="admin-bar">
          <div class="admin-title">
            Knowledge Base Administration
          </div>
          <div class="admin-sub">
            Manage uploaded PDFs and inspect the seeded public HPE references.
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    documents = get_documents()

    uploaded_documents = [
        d
        for d in documents
        if d["source_type"] == "Uploaded PDF"
    ]

    st.markdown(
        "### Upload New PDFs"
    )

    category = st.selectbox(
        "Upload category",
        [
            "General",
            "Compute",
            "Management",
            "HPE GreenLake",
            "Networking",
            "Support",
            "Data Protection",
            "Product",
            "Policies",
            "Procedures",
            "Training",
            "Other",
        ],
        key="admin_upload_category",
    )

    uploaded = st.file_uploader(
        "Select PDF files",
        type=["pdf"],
        accept_multiple_files=True,
        key="admin_files",
    )

    if st.button(
        "Upload & Index Documents",
        type="primary",
        use_container_width=True,
    ):
        if not uploaded:
            st.warning(
                "Select at least one PDF."
            )
        else:
            progress = st.progress(0)

            for index, file in enumerate(uploaded):
                ok, message = add_document(
                    file.name,
                    file.getvalue(),
                    category,
                )

                if ok:
                    st.success(message)
                else:
                    st.warning(message)

                progress.progress(
                    (index + 1) / len(uploaded)
                )

            clear_caches()
            st.rerun()

    st.write("")

    st.markdown(
        "### Uploaded PDF Library"
    )

    if not uploaded_documents:
        st.info(
            "No user-uploaded PDFs yet. "
            "The knowledge base is currently populated with public HPE mock/reference records."
        )

    for document in uploaded_documents:
        with st.container(border=True):
            left, right = st.columns(
                [5.5, 1],
                vertical_alignment="center",
            )

            with left:
                st.markdown(
                    f"**📄 {document['filename']}**"
                )

                st.caption(
                    f"{document['category']} · "
                    f"{document['page_count']} pages · "
                    f"{format_bytes(document['file_size'])} · "
                    f"{document['status']}"
                )

            with right:
                if st.button(
                    "Delete",
                    key=f"delete_{document['id']}",
                ):
                    delete_document(
                        document["id"]
                    )
                    clear_caches()
                    st.rerun()

    st.write("")

    st.markdown(
        "### Public HPE Seed Records"
    )

    st.caption(
        "These are concise paraphrased public-reference records used to make the demo look populated immediately."
    )

    for document in documents:
        if document["source_type"] != "Public HPE Reference":
            continue

        with st.expander(
            f"{document['title']} · {document['category']}"
        ):
            st.write(
                document["description"]
            )

            st.caption(
                f"Keywords: {document['keywords']}"
            )

            if document["source_url"]:
                st.link_button(
                    "Open HPE source ↗",
                    document["source_url"],
                )

    if st.button(
        "Exit Admin Mode",
        use_container_width=True,
    ):
        st.session_state.admin_authenticated = False
        st.session_state.page = "Home"
        st.rerun()


# ============================================================
# ANALYTICS
# ============================================================

def analytics_page():
    documents = get_documents()

    total_pages = sum(
        int(d["page_count"] or 0)
        for d in documents
    )

    uploaded_count = sum(
        d["source_type"] == "Uploaded PDF"
        for d in documents
    )

    public_count = sum(
        d["source_type"] == "Public HPE Reference"
        for d in documents
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "Documents",
        len(documents),
    )

    b.metric(
        "Indexed Pages",
        total_pages,
    )

    c.metric(
        "Public HPE References",
        public_count,
    )

    d.metric(
        "Uploaded PDFs",
        uploaded_count,
    )

    st.write("")

    st.markdown(
        "### Category Distribution"
    )

    category_counts = {}

    for document in documents:
        category = document["category"]
        category_counts[category] = (
            category_counts.get(category, 0) + 1
        )

    for category, count in category_counts.items():
        st.progress(
            count / max(1, len(documents)),
            text=f"{category} — {count}",
        )


# ============================================================
# HELP
# ============================================================

def help_page():
    st.markdown(
        """
        <div class="section-title">
          Help
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")

    st.markdown(
        """
### Searching

Use the hero search, top search bar, or Search page.

The application searches:

- Seeded public HPE reference records
- Uploaded PDF text
- PDF page/chunk content
- Product and technology keywords

### PDF Viewer

When an uploaded PDF is selected:

- The matching page is displayed.
- Search terms are highlighted where possible.
- The left/right arrows change pages.
- The zoom selector changes rendering scale.
- Matching pages are listed on the left.

### Public HPE Seed Data

The initial knowledge base includes concise public-reference records covering:

- HPE ProLiant Compute
- HPE iLO
- HPE GreenLake
- HPE Aruba Networking Central
- HPE Aruba Networking switches
- HPE networking portfolio
- HPE Aruba Networking services
- HPE GreenLake Backup and Recovery
- HPE Compute management

These are **demo/reference records**, not proprietary internal HPE procedures.

### Admin

Select ⚙ and enter the configured `ADMIN_PIN`.

Admin can upload and remove PDFs.
        """
    )


# ============================================================
# ROUTER
# ============================================================

page = st.session_state.page

if page == "Home":
    home_page()

elif page == "Search":
    search_page()

elif page == "Browse All":
    browse_all_page()

elif page == "Categories":
    categories_page()

elif page == "Featured":
    featured_page()

elif page == "Recent":
    recent_page()

elif page == "Upload PDF":
    upload_page()

elif page == "Manage Content":
    manage_page()

elif page == "Admin Login":
    admin_login_page()

elif page == "Analytics":
    analytics_page()

elif page == "Help":
    help_page()

else:
    home_page()
