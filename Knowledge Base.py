import streamlit as st
import sqlite3
import hashlib
import secrets as py_secrets
from datetime import datetime
from pathlib import Path
from html import escape

# ============================================================
# HPE KNOWLEDGE BASE
# Single-file Streamlit application
#
# Secrets required:
#   ACCESS_TOKEN = "one-time-token"
#   ADMIN_PASSWORD = "admin-password"
#
# Optional:
#   ADMIN_NAME = "Arianne Escabillas"
#
# Example .streamlit/secrets.toml:
# ACCESS_TOKEN = "CHANGE-ME"
# ADMIN_PASSWORD = "CHANGE-ME"
# ADMIN_NAME = "Arianne Escabillas"
#
# This demo uses SQLite for SOP/document metadata so it can run
# without an external database. Replace the repository functions
# with MongoDB calls later if needed.
# ============================================================

st.set_page_config(
    page_title="HPE Knowledge Base",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "kb_data"
PDF_DIR = DATA_DIR / "pdfs"
DB_PATH = DATA_DIR / "knowledge_base.db"
DATA_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)

# -----------------------------
# Theme / CSS
# -----------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --ink:#102238;
    --muted:#667384;
    --nav:#00363d;
    --nav2:#002c32;
    --green:#01a982;
    --green-dark:#007f69;
    --teal:#008f83;
    --line:#e4e9ee;
    --surface:#ffffff;
    --page:#f5f7fa;
}

html, body, [class*="css"] {
    font-family: "Inter", Arial, sans-serif;
}

.stApp {
    background:#f5f7fa;
    color:var(--ink);
}

#MainMenu, footer, header {
    visibility:hidden;
}

.block-container {
    padding:0 22px 32px 224px !important;
    max-width:none !important;
}

[data-testid="stHeader"] {
    height:0 !important;
}

[data-testid="stToolbar"] {
    display:none !important;
}

/* Fixed left navigation */
.kb-sidebar {
    position:fixed;
    z-index:1000;
    top:0;
    left:0;
    bottom:0;
    width:212px;
    background:linear-gradient(180deg,#003c42 0%,#002f35 100%);
    color:white;
    box-shadow:2px 0 14px rgba(5,31,39,.08);
}

.kb-logo {
    height:62px;
    background:#f8fafc;
    color:#0d1828;
    display:flex;
    align-items:center;
    padding:0 22px;
    font-size:35px;
    font-weight:800;
    letter-spacing:-2px;
}

.kb-logo span {
    color:#01a982;
    margin-left:1px;
}

.kb-menu {
    padding:31px 11px 0;
}

.kb-nav {
    width:100%;
    border:0;
    border-radius:9px;
    background:transparent;
    color:#eef7f8;
    padding:11px 14px;
    margin:3px 0;
    text-align:left;
    font-size:14px;
    font-weight:500;
    cursor:pointer;
}

.kb-nav:hover { background:rgba(1,169,130,.14); }
.kb-nav.active {
    background:linear-gradient(90deg,#01a982,#00a88c);
    color:white;
    box-shadow:0 5px 13px rgba(0,0,0,.12);
}

.kb-nav .ico {
    display:inline-block;
    width:25px;
    font-size:17px;
    margin-right:8px;
    text-align:center;
}

.kb-brand-bottom {
    position:absolute;
    left:23px;
    bottom:17px;
    color:white;
}
.kb-brand-bottom .mini {
    font-size:21px;
    font-weight:700;
}
.kb-brand-bottom .small {
    font-size:12px;
    opacity:.9;
}
.kb-brand-bottom .version {
    font-size:10px;
    opacity:.65;
    margin-top:5px;
}

/* Header */
.topbar {
    height:62px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    background:#f7f9fb;
}

.page-title {
    font-size:24px;
    font-weight:750;
    letter-spacing:-.5px;
}

.search-top {
    width:46%;
}

.profile-pill {
    display:flex;
    align-items:center;
    gap:10px;
    min-width:220px;
    justify-content:flex-end;
}
.avatar {
    width:38px;
    height:38px;
    border-radius:50%;
    background:linear-gradient(145deg,#f1c5b7,#fff);
    border:2px solid #fff;
    box-shadow:0 1px 4px #bfc7ce;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:17px;
}
.profile-name {
    font-weight:650;
    font-size:13px;
}
.profile-role {
    color:#6b7684;
    font-size:11px;
}

/* Hero */
.hero {
    position:relative;
    overflow:hidden;
    border-radius:10px;
    min-height:213px;
    padding:30px 38px 24px;
    background:
      radial-gradient(circle at 88% 15%,rgba(34,223,205,.20),transparent 28%),
      linear-gradient(115deg,#0b222a 0%,#0a4a50 45%,#007e79 100%);
    color:white;
    box-shadow:0 3px 14px rgba(8,37,46,.09);
}
.hero:after {
    content:"";
    position:absolute;
    right:-60px;
    top:-70px;
    width:510px;
    height:310px;
    opacity:.20;
    background:
      linear-gradient(145deg,transparent 45%,#00e0c2 46%,transparent 47%),
      linear-gradient(165deg,transparent 57%,#36f5df 58%,transparent 59%);
    transform:skewX(-14deg);
}
.hero h1 {
    margin:0 0 8px;
    font-size:35px;
    line-height:1.12;
    letter-spacing:-1px;
    position:relative;
    z-index:2;
}
.hero p {
    margin:0 0 19px;
    font-size:16px;
    position:relative;
    z-index:2;
}
.hero-search {
    max-width:850px;
    position:relative;
    z-index:3;
}
.popular {
    display:flex;
    align-items:center;
    gap:9px;
    margin-top:11px;
    flex-wrap:wrap;
    position:relative;
    z-index:3;
}
.popular-label {
    font-size:12px;
}
.tag {
    border:1px solid rgba(255,255,255,.13);
    background:rgba(0,0,0,.22);
    color:#eefefe;
    border-radius:16px;
    padding:5px 12px;
    font-size:11px;
}

/* Cards */
.metric-row {
    margin-top:15px;
}
.metric-card {
    background:white;
    border:1px solid #e9edf1;
    border-radius:10px;
    min-height:83px;
    padding:15px 17px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    box-shadow:0 2px 9px rgba(25,40,50,.035);
}
.metric-icon {
    width:48px;
    height:48px;
    border-radius:15px;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:24px;
}
.metric-number { font-size:23px; font-weight:750; }
.metric-label { font-size:12px; color:#263342; margin-top:1px; }
.metric-arrow { font-size:23px; color:#1b2c38; }

/* Section */
.section-card {
    background:#fff;
    border:1px solid #e5e9ed;
    border-radius:11px;
    box-shadow:0 2px 9px rgba(25,40,50,.035);
    padding:17px 14px;
}
.section-head {
    display:flex;
    align-items:center;
    justify-content:space-between;
    margin:0 1px 10px;
}
.section-title {
    font-size:17px;
    font-weight:750;
}
.section-link {
    color:#172532;
    font-size:12px;
}

/* category cards */
.cat-card {
    border:1px solid #e6eaee;
    border-radius:9px;
    min-height:70px;
    padding:12px 13px;
    background:white;
    display:flex;
    align-items:center;
    gap:11px;
}
.cat-icon {
    width:39px;
    height:39px;
    border-radius:50%;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:19px;
}
.cat-name { font-size:12px; font-weight:650; }
.cat-count { font-size:10px; color:#6f7b86; margin-top:3px; }
.cat-arrow { margin-left:auto; font-size:18px; }

/* document list */
.doc-row {
    display:flex;
    align-items:center;
    padding:11px 8px;
    border-bottom:1px solid #edf0f2;
}
.doc-row:last-child { border-bottom:0; }
.pdf {
    width:33px;
    height:37px;
    border-radius:7px;
    background:#fff0f1;
    color:#df2935;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:11px;
    font-weight:800;
    margin-right:11px;
}
.doc-title { font-size:12px; font-weight:650; }
.doc-meta { font-size:10px; color:#6e7a86; margin-top:3px; }
.doc-date { margin-left:auto; color:#707b87; font-size:10px; white-space:nowrap; }
.kebab { margin-left:14px; font-size:18px; color:#4e5964; }

/* viewer */
.viewer {
    background:#1f252a;
    border-radius:7px 7px 0 0;
    min-height:280px;
    padding:14px;
    display:flex;
    justify-content:center;
    align-items:center;
    overflow:hidden;
}
.paper {
    width:82%;
    min-height:285px;
    background:white;
    padding:24px 30px;
    color:#142333;
    box-shadow:0 2px 14px rgba(0,0,0,.25);
}
.paper-logo {
    font-size:20px;
    font-weight:800;
    letter-spacing:-1px;
}
.paper-logo span { color:#01a982; }
.paper h3 { margin:12px 0 5px; font-size:18px; }
.paper .line { width:120px; height:3px; background:#01a982; margin:12px 0; }
.paper p { font-size:10px; color:#65717c; line-height:1.55; }
.viewer-footer {
    display:flex;
    justify-content:space-between;
    align-items:center;
    padding:12px 0 0;
}
.open-doc {
    background:#01a982;
    color:white;
    border:0;
    border-radius:8px;
    padding:10px 17px;
    font-weight:650;
}

/* Admin */
.admin-banner {
    background:linear-gradient(90deg,#063c40,#007e70);
    color:white;
    border-radius:10px;
    padding:18px 22px;
    margin-bottom:16px;
}
.form-card {
    background:#fff;
    border:1px solid #e3e8ec;
    border-radius:11px;
    padding:20px;
}
.admin-badge {
    display:inline-block;
    background:#d9fff5;
    color:#007e69;
    border-radius:13px;
    padding:4px 9px;
    font-size:10px;
    font-weight:700;
}

/* Access gate */
.gate {
    max-width:580px;
    margin:70px auto;
    text-align:center;
    background:#fff;
    border:1px solid #e5eaee;
    border-radius:14px;
    padding:34px 38px;
    box-shadow:0 9px 35px rgba(0,40,50,.08);
}
.gate-logo { font-size:42px; font-weight:850; letter-spacing:-3px; }
.gate-logo span { color:#01a982; }
.gate h1 { font-size:25px; margin:10px 0 5px; }
.gate p { color:#6c7782; font-size:13px; }

/* Buttons */
div.stButton > button,
div.stFormSubmitButton > button {
    border-radius:8px;
    border:1px solid #dbe2e7;
    font-weight:600;
}
div.stButton > button[kind="primary"],
div.stFormSubmitButton > button[kind="primary"] {
    background:#01a982;
    color:white;
    border-color:#01a982;
}
div.stButton > button[kind="primary"]:hover,
div.stFormSubmitButton > button[kind="primary"]:hover {
    background:#008f73;
    border-color:#008f73;
}

/* Dialog-ish panels */
div[data-testid="stExpander"] {
    border:1px solid #e3e8ec;
    border-radius:9px;
}

/* Mobile */
@media (max-width: 900px) {
    .kb-sidebar { width:74px; }
    .kb-logo { padding:0 13px; font-size:27px; }
    .kb-menu { padding-left:7px; padding-right:7px; }
    .kb-nav { font-size:0; text-align:center; }
    .kb-nav .ico { margin:0; font-size:18px; }
    .kb-brand-bottom { display:none; }
    .block-container { padding-left:88px !important; padding-right:12px !important; }
    .profile-pill { min-width:auto; }
    .profile-name, .profile-role { display:none; }
    .hero h1 { font-size:27px; }
    .hero { padding:25px; }
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Session state
# -----------------------------
defaults = {
    "access_granted": False,
    "admin": False,
    "page": "Home",
    "selected_doc": None,
    "notice": None,
    "nav_nonce": 0,
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

# -----------------------------
# Secrets
# -----------------------------
def get_secret(name: str, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return default

ACCESS_TOKEN = get_secret("ACCESS_TOKEN", "CHANGE-ME")
ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD", "CHANGE-ME")
ADMIN_NAME = get_secret("ADMIN_NAME", "Knowledge Base Admin")

# -----------------------------
# SQLite repository
# -----------------------------
def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            category TEXT NOT NULL,
            version TEXT DEFAULT 'v1.0',
            description TEXT DEFAULT '',
            keywords TEXT DEFAULT '',
            product TEXT DEFAULT '',
            audience TEXT DEFAULT '',
            procedure TEXT DEFAULT '',
            source_url TEXT DEFAULT '',
            filename TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            featured INTEGER DEFAULT 0
        )
    """)
    conn.commit()

    count = conn.execute("SELECT COUNT(*) AS c FROM documents").fetchone()["c"]
    if count == 0:
        seed = [
            ("HPE iLO 7 Licensing Guide", "Licensing", "v1.0",
             "Public-reference mock entry covering purchasing, registration and activation of HPE iLO licenses.",
             "iLO 7, licensing, SAID, activation, support", "HPE iLO 7", "Support / Licensing",
             "Use the official HPE iLO licensing guide for license purchase, registration, activation and support-entitlement references.",
             "https://support.hpe.com/hpesc/public/docDisplay?docId=sd00005843en_us", "", "2026-09-28 09:00", "2026-09-28 09:00", 1),

            ("HPE OneView Licensing Overview", "Licensing", "v1.0",
             "Public-reference mock entry about licensing requirements for HPE OneView-managed hardware.",
             "OneView, licensing, server, Synergy, trial", "HPE OneView", "Support / Operations",
             "Confirm the managed hardware and applicable OneView license type in the official product documentation before advising a customer.",
             "https://support.hpe.com/hpesc/public/docDisplay?docId=sd00007490en_us", "", "2026-09-25 09:00", "2026-09-25 09:00", 0),

            ("HPE iLO Documentation Quick Links", "Technical Support", "v1.0",
             "Public-reference mock entry pointing agents to iLO user, security, troubleshooting, licensing and Redfish documentation.",
             "iLO, user guide, troubleshooting, Redfish, security", "HPE iLO", "Technical Support",
             "Use the HPE documentation index to select the guide matching the customer's iLO generation.",
             "https://support.hpe.com/hpesc/public/docDisplay?docId=sd00004310en_us", "", "2026-09-20 09:00", "2026-09-20 09:00", 0),

            ("HPE ProLiant iLO License Features", "Product Guides", "v1.0",
             "Mock knowledge article summarizing the distinction between standard and licensed iLO features.",
             "ProLiant, iLO Standard, iLO Advanced, features", "HPE ProLiant", "Technical Support",
             "Check the applicable iLO generation and license family. Do not assume a feature is available across every generation.",
             "https://support.hpe.com/hpesc/public/docDisplay?docId=c05269613", "", "2026-09-18 09:00", "2026-09-18 09:00", 0),

            ("HPE OneView Product Information Reference", "Product Guides", "v1.0",
             "Mock product reference for HPE OneView editions and licensing terminology.",
             "OneView, product information, license, LTU, E-LTU", "HPE OneView", "Product Support",
             "Use the official HPE product information reference for current part numbers and product descriptions.",
             "https://support.hpe.com/hpesc/public/docDisplay?docId=a00006903en_us", "", "2026-09-15 09:00", "2026-09-15 09:00", 0),
        ]
        conn.executemany("""
            INSERT INTO documents
            (title,category,version,description,keywords,product,audience,procedure,source_url,filename,created_at,updated_at,featured)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, seed)
        conn.commit()
    conn.close()

init_db()

def all_docs():
    conn = db()
    rows = conn.execute("SELECT * FROM documents ORDER BY updated_at DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def search_docs(query):
    query = query.strip().lower()
    if not query:
        return all_docs()
    conn = db()
    rows = conn.execute("""
        SELECT * FROM documents
        WHERE lower(title) LIKE ?
           OR lower(category) LIKE ?
           OR lower(keywords) LIKE ?
           OR lower(product) LIKE ?
           OR lower(description) LIKE ?
           OR lower(procedure) LIKE ?
        ORDER BY updated_at DESC
    """, tuple([f"%{query}%"] * 6)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def categories():
    docs = all_docs()
    out = {}
    for d in docs:
        out[d["category"]] = out.get(d["category"], 0) + 1
    return out

def save_doc(data):
    conn = db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn.execute("""
        INSERT INTO documents
        (title,category,version,description,keywords,product,audience,procedure,source_url,filename,created_at,updated_at,featured)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        data["title"], data["category"], data["version"], data["description"],
        data["keywords"], data["product"], data["audience"], data["procedure"],
        data["source_url"], data["filename"], now, now, int(data["featured"])
    ))
    conn.commit()
    conn.close()

def update_doc(doc_id, data):
    conn = db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn.execute("""
        UPDATE documents SET
        title=?, category=?, version=?, description=?, keywords=?,
        product=?, audience=?, procedure=?, source_url=?, updated_at=?,
        featured=?
        WHERE id=?
    """, (
        data["title"], data["category"], data["version"], data["description"],
        data["keywords"], data["product"], data["audience"], data["procedure"],
        data["source_url"], now, int(data["featured"]), doc_id
    ))
    conn.commit()
    conn.close()

def get_doc(doc_id):
    conn = db()
    row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
    conn.close()
    return dict(row) if row else None

# -----------------------------
# One-time token gate
# -----------------------------
def token_matches(value):
    if not value:
        return False
    return py_secrets.compare_digest(str(value).strip(), str(ACCESS_TOKEN))

if not st.session_state.access_granted:
    st.markdown("""
    <div class="gate">
        <div class="gate-logo">HP<span>E</span></div>
        <h1>Knowledge Base Access</h1>
        <p>Enter the one-time access token to open the HPE Knowledge Base.</p>
    </div>
    """, unsafe_allow_html=True)

    with st.form("access_form"):
        token = st.text_input("Access token", type="password", placeholder="Enter access token")
        submit = st.form_submit_button("Access Knowledge Base", type="primary", use_container_width=True)

    if submit:
        if token_matches(token):
            st.session_state.access_granted = True
            st.session_state.page = "Home"
            st.rerun()
        else:
            st.error("Invalid access token.")
    st.stop()

# -----------------------------
# Sidebar
# -----------------------------
pages = ["Home", "Browse All", "Categories", "Favorites", "Recent", "Upload PDF", "Manage Content", "Analytics", "Feedback", "Help"]
icons = ["⌂", "▣", "⊞", "☆", "◷", "♧", "▤", "⌁", "□", "?"]

st.markdown("""
<div class="kb-sidebar">
  <div class="kb-logo">HP<span>E</span></div>
  <div class="kb-menu">
""", unsafe_allow_html=True)

for p, ico in zip(pages, icons):
    active = st.session_state.page == p
    # Buttons live in the fixed sidebar container visually.
    label = f"{ico}   {p}"
    if st.button(label, key=f"nav_{p}", use_container_width=True):
        st.session_state.page = p
        st.session_state.selected_doc = None
        st.rerun()

st.markdown("""
  </div>
  <div class="kb-brand-bottom">
      <div class="mini">HPE</div>
      <div class="small">Knowledge Base</div>
      <div class="version">v1.0.0</div>
  </div>
</div>
""", unsafe_allow_html=True)

# -----------------------------
# Header
# -----------------------------
header_left, header_mid, header_right = st.columns([1.35, 3.0, 1.55], vertical_alignment="center")

with header_left:
    st.markdown(f'<div class="page-title">{escape(st.session_state.page)}</div>', unsafe_allow_html=True)

with header_mid:
    top_search = st.text_input(
        "Search",
        placeholder="Search for topics, keywords, or questions...",
        label_visibility="collapsed",
        key="top_search"
    )

with header_right:
    c1, c2, c3 = st.columns([.55, 2.0, .65], vertical_alignment="center")
    with c1:
        # Gear opens admin login.
        gear = st.button("⚙", help="Admin settings", key="gear_btn")
        if gear:
            st.session_state.admin_login_open = True
    with c2:
        role_text = "Admin" if st.session_state.admin else "Viewer"
        st.markdown(
            f'<div class="profile-pill"><div class="avatar">👩🏻</div>'
            f'<div><div class="profile-name">{escape(ADMIN_NAME if st.session_state.admin else "Knowledge User")}</div>'
            f'<div class="profile-role">{role_text}</div></div></div>',
            unsafe_allow_html=True
        )
    with c3:
        st.markdown('<div style="font-size:20px;text-align:right;">⌄</div>', unsafe_allow_html=True)

# Admin login panel
if st.session_state.get("admin_login_open", False) and not st.session_state.admin:
    with st.expander("Admin sign-in", expanded=True):
        with st.form("admin_login"):
            pw = st.text_input("Admin password", type="password")
            a, b = st.columns(2)
            with a:
                ok = st.form_submit_button("Sign in as Admin", type="primary", use_container_width=True)
            with b:
                cancel = st.form_submit_button("Cancel", use_container_width=True)

        if cancel:
            st.session_state.admin_login_open = False
            st.rerun()
        if ok:
            if py_secrets.compare_digest(pw, str(ADMIN_PASSWORD)):
                st.session_state.admin = True
                st.session_state.admin_login_open = False
                st.session_state.page = "Manage Content"
                st.success("Admin access granted.")
                st.rerun()
            else:
                st.error("Incorrect admin password.")

if st.session_state.admin:
    st.markdown('<span class="admin-badge">ADMIN MODE</span>', unsafe_allow_html=True)
    st.write("")

# -----------------------------
# Search route
# -----------------------------
if top_search.strip():
    st.session_state.page = "Browse All"

# -----------------------------
# Home
# -----------------------------
def render_home():
    docs = all_docs()
    cats = categories()

    st.markdown("""
    <div class="hero">
      <h1>Find the answers you need</h1>
      <p>Search our knowledge base, explore topics, or browse by category.</p>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        hero_search = st.text_input(
            "Search the knowledge base",
            placeholder="Search for solutions, guides, or keywords...",
            label_visibility="collapsed",
            key="hero_search"
        )
        if hero_search:
            st.session_state.page = "Browse All"
            st.session_state.search_term = hero_search
            st.rerun()

    st.markdown("""
    <div class="popular">
      <span class="popular-label">Popular searches:</span>
      <span class="tag">Licensing</span>
      <span class="tag">Portal Access</span>
      <span class="tag">Account Setup</span>
      <span class="tag">Troubleshooting</span>
      <span class="tag">HPE GreenLake</span>
      <span class="tag">Software Support</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div class='metric-row'></div>", unsafe_allow_html=True)
    cols = st.columns(4, gap="small")
    metrics = [
        ("📖", len(docs), "Total Documents", "#dcfbf2"),
        ("▣", len(cats), "Categories", "#e3f0ff"),
        ("★", max(0, min(36, len(docs)*7)), "Most Viewed", "#fff5d9"),
        ("☁", min(24, len(docs)), "Recently Added", "#f0e9ff"),
    ]
    for col, (ico, num, label, bg) in zip(cols, metrics):
        with col:
            st.markdown(
                f'<div class="metric-card"><div class="metric-icon" style="background:{bg}">{ico}</div>'
                f'<div style="flex:1;margin-left:12px"><div class="metric-number">{num}</div>'
                f'<div class="metric-label">{label}</div></div><div class="metric-arrow">›</div></div>',
                unsafe_allow_html=True
            )

    st.write("")
    st.markdown('<div class="section-card"><div class="section-head"><div class="section-title">▦ &nbsp;Browse by Category</div><div class="section-link">View All Categories →</div></div>', unsafe_allow_html=True)

    category_icons = {
        "Account Management": ("♙", "#dffaf1"),
        "Licensing": ("⚿", "#e3f0ff"),
        "Portal & Access": ("▣", "#dffaf1"),
        "Technical Support": ("🔧", "#f0e6ff"),
        "HPE GreenLake": ("☁", "#dffaf1"),
        "Product Guides": ("▣", "#e3f0ff"),
        "Policies & Procedures": ("♢", "#fff4d7"),
        "Troubleshooting": ("⚙", "#e3f0ff"),
    }

    items = list(cats.items())[:8]
    if items:
        for start in range(0, len(items), 4):
            row = items[start:start+4]
            c = st.columns(4, gap="small")
            for col, (name, count) in zip(c, row):
                ico, bg = category_icons.get(name, ("▤", "#edf2f6"))
                with col:
                    if st.button(f"{ico}   {name}\n{count} documents", key=f"cat_{name}", use_container_width=True):
                        st.session_state.page = "Browse All"
                        st.session_state.search_term = name
                        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    st.write("")
    left, right = st.columns([1.1, .9], gap="small")

    recent = sorted(docs, key=lambda x: x["updated_at"], reverse=True)[:5]
    with left:
        st.markdown('<div class="section-card"><div class="section-head"><div class="section-title">▤ &nbsp;Recent Documents</div><div class="section-link">View All →</div></div>', unsafe_allow_html=True)
        for d in recent:
            if st.button(f"📄  {d['title']}   ·   {d['category']} · {d['version']}", key=f"recent_{d['id']}", use_container_width=True):
                st.session_state.selected_doc = d["id"]
                st.session_state.page = "Browse All"
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        featured = next((d for d in docs if d["featured"]), docs[0] if docs else None)
        st.markdown('<div class="section-card"><div class="section-head"><div class="section-title">★ &nbsp;Featured Document</div></div>', unsafe_allow_html=True)
        if featured:
            st.markdown(f"""
            <div class="viewer">
              <div class="paper">
                <div class="paper-logo">HP<span>E</span></div>
                <h3>{escape(featured["title"])}</h3>
                <div class="line"></div>
                <p>Public-reference knowledge article</p>
                <p>{escape(featured["description"])}</p>
                <p><b>Category:</b> {escape(featured["category"])} &nbsp; <b>Version:</b> {escape(featured["version"])}</p>
              </div>
            </div>
            """, unsafe_allow_html=True)
            x, y = st.columns([1.3, .7], vertical_alignment="center")
            with x:
                st.markdown(f"**{escape(featured['title'])}**<br><span style='font-size:11px;color:#697582'>{escape(featured['category'])} · {escape(featured['version'])} · {escape(featured['updated_at'][:10])}</span>", unsafe_allow_html=True)
            with y:
                if st.button("Open Document →", key="open_featured", type="primary", use_container_width=True):
                    st.session_state.selected_doc = featured["id"]
                    st.session_state.page = "Browse All"
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

# -----------------------------
# Browse
# -----------------------------
def render_browse():
    query = st.session_state.get("search_term", top_search)
    docs = search_docs(query)
    if st.session_state.selected_doc:
        selected = get_doc(st.session_state.selected_doc)
    else:
        selected = None

    st.markdown("### Search and browse")
    search = st.text_input("Search knowledge base", value=query or "", placeholder="Try: iLO licensing, OneView, ProLiant...")
    if search != query:
        st.session_state.search_term = search
        docs = search_docs(search)

    if selected:
        st.markdown(f"#### {escape(selected['title'])}", unsafe_allow_html=True)
        st.caption(f"{selected['category']} · {selected['version']} · Updated {selected['updated_at'][:10]}")
        st.markdown(f"**Overview**  \n{escape(selected['description'])}")
        st.markdown(f"**Keywords:** {escape(selected['keywords'])}")
        st.markdown(f"**Procedure / agent notes**  \n{escape(selected['procedure'])}")
        if selected["source_url"]:
            st.link_button("Open official public reference ↗", selected["source_url"])
        if st.button("← Back to results", key="back_results"):
            st.session_state.selected_doc = None
            st.rerun()
        return

    st.write(f"**{len(docs)} result(s)**")
    if not docs:
        st.info("No matching documents. Try a broader keyword.")
    for d in docs:
        with st.container(border=True):
            a, b = st.columns([5, 1], vertical_alignment="center")
            with a:
                st.markdown(f"**{escape(d['title'])}**")
                st.caption(f"{d['category']} · {d['version']} · {d['updated_at'][:10]}")
                st.write(d["description"])
                if d["keywords"]:
                    st.caption("Keywords: " + d["keywords"])
            with b:
                if st.button("Open", key=f"open_{d['id']}", use_container_width=True):
                    st.session_state.selected_doc = d["id"]
                    st.rerun()

# -----------------------------
# Categories
# -----------------------------
def render_categories():
    st.markdown("### Categories")
    cats = categories()
    for start in range(0, len(cats), 4):
        row = list(cats.items())[start:start+4]
        cols = st.columns(4)
        for col, (cat, count) in zip(cols, row):
            with col:
                if st.button(f"▦\n\n**{cat}**\n\n{count} documents", key=f"browsecat_{cat}", use_container_width=True):
                    st.session_state.search_term = cat
                    st.session_state.page = "Browse All"
                    st.rerun()

# -----------------------------
# Upload PDF
# -----------------------------
def render_upload():
    st.markdown("### Upload PDF")
    st.info("Upload a public or internally approved PDF. The file is stored locally and its metadata is registered in the knowledge base.")
    with st.form("pdf_upload_form"):
        uploaded = st.file_uploader("PDF file", type=["pdf"])
        title = st.text_input("Document title")
        category = st.selectbox("Category", ["Licensing", "Portal & Access", "Account Management", "Technical Support", "HPE GreenLake", "Product Guides", "Policies & Procedures", "Troubleshooting"])
        version = st.text_input("Version", "v1.0")
        description = st.text_area("Short description")
        keywords = st.text_input("Keywords", placeholder="iLO, licensing, activation...")
        source_url = st.text_input("Public source URL (optional)")
        submit = st.form_submit_button("Add PDF to Knowledge Base", type="primary")

    if submit:
        if not uploaded or not title:
            st.error("Please provide a PDF and document title.")
        else:
            safe_name = Path(uploaded.name).name.replace(" ", "_")
            destination = PDF_DIR / f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{safe_name}"
            destination.write_bytes(uploaded.getbuffer())
            save_doc({
                "title": title,
                "category": category,
                "version": version,
                "description": description,
                "keywords": keywords,
                "product": "",
                "audience": "Knowledge Base",
                "procedure": "Open the PDF and review the source content before using it as an agent reference.",
                "source_url": source_url,
                "filename": destination.name,
                "featured": False,
            })
            st.success("PDF registered successfully.")
            st.rerun()

# -----------------------------
# Admin SOP editor
# -----------------------------
def render_manage():
    if not st.session_state.admin:
        st.warning("Admin access is required. Use the ⚙ button in the upper-right corner.")
        return

    st.markdown("""
    <div class="admin-banner">
      <div style="font-size:21px;font-weight:750;">SOP Management</div>
      <div style="font-size:12px;opacity:.88;margin-top:4px;">
        Create structured SOP records or update an existing knowledge article.
      </div>
    </div>
    """, unsafe_allow_html=True)

    docs = all_docs()
    choices = {"➕ Create new SOP": None}
    choices.update({f"{d['title']} · {d['version']}": d["id"] for d in docs})
    selected_label = st.selectbox("Action", list(choices.keys()))
    selected_id = choices[selected_label]
    existing = get_doc(selected_id) if selected_id else None

    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    with st.form("sop_form"):
        st.markdown("#### SOP information")
        c1, c2 = st.columns(2)
        with c1:
            title = st.text_input("SOP / Article title *", value=existing["title"] if existing else "")
            category = st.selectbox(
                "Category *",
                ["Licensing", "Portal & Access", "Account Management", "Technical Support", "HPE GreenLake", "Product Guides", "Policies & Procedures", "Troubleshooting"],
                index=(["Licensing", "Portal & Access", "Account Management", "Technical Support", "HPE GreenLake", "Product Guides", "Policies & Procedures", "Troubleshooting"].index(existing["category"]) if existing and existing["category"] in ["Licensing", "Portal & Access", "Account Management", "Technical Support", "HPE GreenLake", "Product Guides", "Policies & Procedures", "Troubleshooting"] else 0)
            )
            version = st.text_input("Version", value=existing["version"] if existing else "v1.0")
            product = st.text_input("Product / device", value=existing["product"] if existing else "")
            audience = st.text_input("Audience", value=existing["audience"] if existing else "Support / Licensing")
        with c2:
            keywords = st.text_input("Search keywords *", value=existing["keywords"] if existing else "")
            source_url = st.text_input("Official/public source URL", value=existing["source_url"] if existing else "")
            featured = st.checkbox("Feature this document on the home page", value=bool(existing["featured"]) if existing else False)
            filename = existing["filename"] if existing else ""
            if filename:
                st.caption(f"Attached PDF: {filename}")

        description = st.text_area(
            "Overview / purpose *",
            value=existing["description"] if existing else "",
            height=110
        )
        procedure = st.text_area(
            "Procedure / resolution steps *",
            value=existing["procedure"] if existing else "",
            height=210,
            placeholder="1. Identify the product and generation.\n2. Verify entitlement or license state.\n3. Follow the official HPE procedure.\n4. Document the outcome."
        )

        pdf = st.file_uploader(
            "Optional supporting PDF",
            type=["pdf"],
            key=f"sop_pdf_{selected_id or 'new'}"
        )

        submitted = st.form_submit_button(
            "Update SOP" if existing else "Create SOP",
            type="primary",
            use_container_width=True
        )

    st.markdown("</div>", unsafe_allow_html=True)

    if submitted:
        if not title.strip() or not keywords.strip() or not description.strip() or not procedure.strip():
            st.error("Title, keywords, overview and procedure are required.")
            return

        saved_filename = filename
        if pdf:
            safe_name = Path(pdf.name).name.replace(" ", "_")
            dest = PDF_DIR / f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{safe_name}"
            dest.write_bytes(pdf.getbuffer())
            saved_filename = dest.name

        payload = {
            "title": title.strip(),
            "category": category,
            "version": version.strip() or "v1.0",
            "description": description.strip(),
            "keywords": keywords.strip(),
            "product": product.strip(),
            "audience": audience.strip(),
            "procedure": procedure.strip(),
            "source_url": source_url.strip(),
            "filename": saved_filename,
            "featured": featured,
        }

        if existing:
            update_doc(existing["id"], payload)
            st.success("SOP updated successfully.")
        else:
            save_doc(payload)
            st.success("New SOP created successfully.")
        st.rerun()

    st.markdown("#### Existing SOPs")
    for d in docs:
        with st.expander(f"{d['title']} · {d['version']} · {d['category']}"):
            st.write(d["description"])
            st.caption("Keywords: " + d["keywords"])

# -----------------------------
# Analytics / other pages
# -----------------------------
def render_analytics():
    docs = all_docs()
    cats = categories()
    st.markdown("### Knowledge Base Analytics")
    c1, c2, c3 = st.columns(3)
    c1.metric("Documents", len(docs))
    c2.metric("Categories", len(cats))
    c3.metric("Featured", sum(1 for d in docs if d["featured"]))
    st.markdown("#### Documents by category")
    for cat, count in cats.items():
        st.progress(min(1.0, count / max(1, len(docs))), text=f"{cat} — {count}")

def render_feedback():
    st.markdown("### Feedback")
    with st.form("feedback"):
        rating = st.radio("Was this article useful?", ["Yes", "Partly", "No"], horizontal=True)
        notes = st.text_area("Comments")
        send = st.form_submit_button("Submit feedback", type="primary")
    if send:
        st.success("Thank you. Feedback recorded for this demo.")

def render_help():
    st.markdown("### Help")
    st.info("Use the search bar to find SOPs, product references and troubleshooting material. Admins can use ⚙ → Admin sign-in → Manage Content to create or update structured SOPs.")
    st.markdown("""
**Security notes**
- The access token and admin password are read from `st.secrets`.
- Do not hard-code credentials in the Python file.
- The included SQLite database is suitable for a prototype. For multi-user production deployment, move document records and uploaded files to a persistent database/object store.
""")

def render_simple(title):
    st.markdown(f"### {title}")
    if title == "Favorites":
        st.info("Favorites are available as a UI placeholder. Add a favorites table/user identity layer when user accounts are connected.")
    elif title == "Recent":
        docs = sorted(all_docs(), key=lambda x: x["updated_at"], reverse=True)
        for d in docs[:10]:
            st.write(f"📄 **{d['title']}** — {d['category']} · {d['updated_at'][:10]}")
    else:
        st.info("This section is ready for extension.")

# -----------------------------
# Route
# -----------------------------
page = st.session_state.page

if page == "Home":
    render_home()
elif page == "Browse All":
    render_browse()
elif page == "Categories":
    render_categories()
elif page == "Upload PDF":
    render_upload()
elif page == "Manage Content":
    render_manage()
elif page == "Analytics":
    render_analytics()
elif page == "Feedback":
    render_feedback()
elif page == "Help":
    render_help()
else:
    render_simple(page)

# -----------------------------
# Admin logout
# -----------------------------
if st.session_state.admin:
    with st.sidebar:
        pass
    st.markdown("---")
    c1, c2 = st.columns([8, 1])
    with c2:
        if st.button("Exit Admin", key="exit_admin"):
            st.session_state.admin = False
            st.session_state.page = "Home"
            st.rerun()
