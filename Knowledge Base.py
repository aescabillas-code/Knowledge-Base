import os
import re
import io
import json
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

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ============================================================
# CONFIGURATION & STORAGE
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
    initial_sidebar_state="collapsed",
)

# ============================================================
# UI STYLING & SYSTEM CSS
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
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}

.stApp {
    background: linear-gradient(180deg, #f8fbfc 0%, #f2f6f8 100%);
    color: var(--text);
}

[data-testid="stHeader"] { 
    height: 0 !important; 
    min-height: 0 !important; 
    background: transparent; 
}

[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"],
[data-testid="stMainMenu"],
button[kind="header"],
[data-testid="stHeaderActionElements"],
[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"] {
    display: none !important;
    visibility: hidden !important;
}

/* Unified Topbar Container */
.st-key-topbar_container {
    border-bottom: 1px solid #dfe7eb;
    background: linear-gradient(105deg, #ffffff 0%, #f7fbfc 70%, #e8f7f7 100%);
    border-radius: 8px;
    padding: 10px 18px;
    margin-bottom: 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.03);
}

.brand-mark {
    width: 42px;
    height: 6px;
    background-color: var(--teal);
    margin-bottom: 6px;
    border-radius: 2px;
}

.top-user-pill {
    background: #e8f7f7;
    color: #315468;
    font-size: 12px;
    padding: 6px 14px;
    border-radius: 20px;
    font-weight: 600;
    white-space: nowrap;
    border: 1px solid #cce5df;
    display: inline-flex;
    align-items: center;
    height: 34px;
}

/* Gear Button Restyling in Header */
.st-key-topbar_container div[data-testid="stPopover"] > button {
    border-radius: 6px !important;
    border: 1px solid #cce5df !important;
    background: #ffffff !important;
    color: #315468 !important;
    height: 34px !important;
    min-height: 34px !important;
    padding: 0 10px !important;
    font-size: 15px !important;
}

.st-key-topbar_container div[data-testid="stPopover"] > button:hover {
    background: #e8f7f7 !important;
    border-color: var(--teal) !important;
    color: var(--teal) !important;
}

/* Direct Extracted Answer Card */
.exact-answer-card {
    background: linear-gradient(110deg, #f0fdf4 0%, #ffffff 85%);
    border: 1px solid #86efac;
    border-left: 5px solid var(--teal);
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 16px;
    box-shadow: 0 2px 8px rgba(0, 169, 130, 0.05);
}
.exact-answer-head { display: flex; justify-content: space-between; align-items: center; }
.exact-answer-title { color: #047857; font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
.match-pill { background: #dcfce7; color: #15803d; font-weight: 700; padding: 3px 10px; border-radius: 12px; font-size: 11px; }
.exact-answer-text {
    margin: 8px 0;
    color: #0f172a;
    font-size: 15px;
    line-height: 1.5;
    font-weight: 500;
}
.answer-meta { font-size: 11px; color: #64748b; }

/* Document Reader Frame */
.reader-toolbar {
    background: #ffffff;
    border: 1px solid var(--border);
    border-radius: 8px 8px 0 0;
    padding: 10px 14px;
    border-bottom: 0;
}
.reader-title { font-size: 14px; font-weight: 700; color: var(--text); word-break: break-all; }
.reader-meta { font-size: 11px; color: var(--muted); margin-top: 2px; }

/* Evidence Result Cards */
.result-card {
    background: #ffffff;
    border: 1px solid var(--border);
    border-radius: 7px;
    padding: 10px 12px;
    margin-bottom: 8px;
}
.result-label { font-weight: 700; font-size: 11px; text-transform: uppercase; }
.result-filename { color: #0369a1; font-weight: 700; font-size: 12.5px; margin-top: 2px; word-break: break-word; }
.result-snippet { color: #475569; font-size: 11.5px; line-height: 1.45; margin: 6px 0; }

/* Authentication Gate */
.auth-shell { max-width: 520px; margin: 80px auto 20px; text-align: center; }
.auth-brand-mark { width: 48px; height: 6px; background-color: var(--teal); margin: 0 auto 10px; border-radius: 2px; }
.auth-brand { font-size: 13px; font-weight: 700; color: #111; }
.auth-title { margin-top: 18px; font-size: 28px; font-weight: 700; color: var(--text); }
.auth-subtitle { margin-top: 6px; color: var(--muted); font-size: 13px; margin-bottom: 24px; }

/* Welcome Screen Card */
.welcome-card {
    margin: 40px auto; max-width: 680px; text-align: center; background: white;
    border: 1px solid var(--border); border-radius: 12px; padding: 36px;
    box-shadow: 0 4px 16px rgba(12, 54, 70, 0.04);
}
.welcome-icon { font-size: 38px; color: var(--teal); }
.welcome-title { font-size: 22px; font-weight: 700; color: var(--text); margin-top: 8px; }
.welcome-text { color: var(--muted); max-width: 500px; margin: 8px auto; line-height: 1.5; font-size: 13.5px; }
.welcome-stats { display: flex; justify-content: center; gap: 28px; color: #57707e; margin-top: 16px; font-size: 12px; }

button[kind="primary"] { background: var(--teal) !important; border-color: var(--teal) !important; }
button[kind="primary"]:hover { background: var(--teal-dark) !important; }
</style>
""",
    unsafe_allow_html=True,
)

# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = db()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
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
# PERSISTENT BROWSER ACCESS GATE
# ============================================================

ACCESS_CODE = str(st.secrets.get("ACCESS_CODE", os.getenv("ACCESS_CODE", ""))).strip()
TOKEN_SECRET = str(st.secrets.get("TOKEN_SECRET", os.getenv("TOKEN_SECRET", ""))).strip()

if not TOKEN_SECRET:
    TOKEN_SECRET = hashlib.sha256(f"{os.getcwd()}::{os.getenv('HOSTNAME', 'kb_app')}".encode()).hexdigest()

def get_token_serializer():
    if URLSafeTimedSerializer is None or not TOKEN_SECRET:
        return None
    return URLSafeTimedSerializer(TOKEN_SECRET, salt="knowledge-base-browser-access")

def create_browser_token():
    serializer = get_token_serializer()
    return serializer.dumps({"authorized": True}) if serializer else ""

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

def browser_is_authorized():
    if st.session_state.get("access_authorized"):
        return True
    token = st.query_params.get("kb_access", "")
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

def get_admin_pin():
    return str(st.secrets.get("ADMIN_PIN", os.getenv("ADMIN_PIN", ""))).strip()

def render_access_gate():
    st.markdown(
        """
        <div class="auth-shell">
            <div class="auth-brand-mark"></div>
            <div class="auth-brand">HEWLETT PACKARD ENTERPRISE</div>
            <div class="auth-title">Knowledge Base</div>
            <div class="auth-subtitle">Authorized verification engine for organizational documents.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not ACCESS_CODE:
        st.error("System access gate is unconfigured. Set ACCESS_CODE in Streamlit Secrets or Environment.")
        st.stop()

    _, center, _ = st.columns([1, 1.4, 1])
    with center:
        with st.form("access_code_form"):
            entered_code = st.text_input("Access Code", type="password", placeholder="Enter authorization key")
            submitted = st.form_submit_button("Authenticate Workspace", type="primary", use_container_width=True)

        if submitted:
            if entered_code.strip() == ACCESS_CODE:
                if authorize_browser():
                    st.rerun()
                else:
                    st.error("Crypto subsystem error: Unable to issue token.")
            else:
                st.error("Invalid access code.")

if not browser_is_authorized():
    render_access_gate()
    st.stop()

# ============================================================
# STATE INITIALIZATION
# ============================================================

for key, default in [
    ("page", "Search"),
    ("selected_document", None),
    ("selected_page", 1),
    ("search_query", ""),
    ("admin_authenticated", False),
    ("force_result_id", None),
    ("search_results", []),
    ("search_signature", None),
    ("viewer_page", None),
    ("highlight_target", ""),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ============================================================
# UTILITIES & REPOSITORY
# ============================================================

def clean_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def make_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def split_text(text: str, chunk_size=900, overlap=150):
    words = text.split()
    if not words:
        return []
    chunks, start = [], 0
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

def add_document(file_name, pdf_bytes, category="General"):
    file_hash = make_hash(pdf_bytes)
    conn = db()
    exists = conn.execute("SELECT id FROM documents WHERE file_hash = ?", (file_hash,)).fetchone()
    if exists:
        conn.close()
        return False, "This document already exists in the knowledge base."

    try:
        pages, page_count = extract_pdf(pdf_bytes)
    except Exception as e:
        conn.close()
        return False, f"Corrupted or invalid PDF: {e}"

    stored_path = save_pdf(file_name, pdf_bytes, file_hash)
    now = datetime.now().isoformat(timespec="seconds")

    cursor = conn.execute(
        """
        INSERT INTO documents (filename, stored_path, file_hash, category, page_count, file_size, uploaded_at, indexed_at, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (file_name, stored_path, file_hash, category, page_count, len(pdf_bytes), now, now, "Indexed"),
    )
    doc_id = cursor.lastrowid

    for page_number, page_text in pages:
        for chunk_idx, chunk in enumerate(split_text(page_text)):
            conn.execute(
                "INSERT INTO chunks (document_id, page_number, chunk_index, text) VALUES (?, ?, ?, ?)",
                (doc_id, page_number, chunk_idx, chunk),
            )
    conn.commit()
    conn.close()
    return True, f"Indexed {file_name} ({page_count} pages)."

@st.cache_data(ttl=60, show_spinner=False)
def get_documents():
    conn = db()
    rows = conn.execute("SELECT * FROM documents ORDER BY uploaded_at DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]

@st.cache_data(ttl=60, show_spinner=False)
def get_categories():
    conn = db()
    rows = conn.execute("SELECT DISTINCT category FROM documents ORDER BY category").fetchall()
    conn.close()
    return [r["category"] for r in rows]

def get_all_chunks():
    conn = db()
    rows = conn.execute(
        """
        SELECT c.id, c.document_id, c.page_number, c.chunk_index, c.text, d.filename, d.category, d.stored_path
        FROM chunks c JOIN documents d ON d.id = c.document_id ORDER BY c.id
        """
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def delete_document(document_id):
    conn = db()
    row = conn.execute("SELECT stored_path FROM documents WHERE id = ?", (document_id,)).fetchone()
    if row:
        Path(row["stored_path"]).unlink(missing_ok=True)
    conn.execute("DELETE FROM chunks WHERE document_id = ?", (document_id,))
    conn.execute("DELETE FROM documents WHERE id = ?", (document_id,))
    conn.commit()
    conn.close()

def format_bytes(val):
    if not val:
        return "0 B"
    val = float(val)
    for unit in ["B", "KB", "MB", "GB"]:
        if val < 1024:
            return f"{val:.1f} {unit}"
        val /= 1024
    return f"{val:.1f} TB"

def extract_direct_sentence(text: str, query: str) -> str:
    """Isolate the single most accurate, verbatim sentence from a matching passage."""
    sentences = re.split(r"(?<=[.!?])\s+", clean_text(text))
    terms = [t.lower() for t in re.findall(r"\w+", query) if len(t) > 2]
    if not terms or not sentences:
        return text[:280] + "..." if len(text) > 280 else text

    best_sent = sentences[0]
    best_score = -1
    for s in sentences:
        low = s.lower()
        score = sum(low.count(term) for term in terms)
        if score > best_score and len(s.split()) > 3:
            best_score = score
            best_sent = s.strip()

    return best_sent

def make_snippet(text: str, query: str, radius=180) -> str:
    clean = " ".join(text.split())
    terms = [t.lower() for t in re.findall(r"\w+", query) if len(t) > 2]
    if not terms:
        return clean[:radius] + ("..." if len(clean) > radius else "")

    low = clean.lower()
    matches = [low.find(t) for t in terms if low.find(t) >= 0]
    if not matches:
        return clean[:radius] + ("..." if len(clean) > radius else "")

    start = max(0, min(matches) - 50)
    end = min(len(clean), start + radius)
    return ("..." if start > 0 else "") + clean[start:end] + ("..." if end < len(clean) else "")

# ============================================================
# SEARCH & HIGHLIGHT RENDERING ENGINE
# ============================================================

@st.cache_data(ttl=60, show_spinner=False)
def build_search_index():
    rows = get_all_chunks()
    if not rows:
        return None, None, []
    texts = [r["text"] for r in rows]
    vectorizer = TfidfVectorizer(
        lowercase=True,
        stop_words="english",
        ngram_range=(1, 2),
        sublinear_tf=True,
        dtype="float32",
    )
    matrix = vectorizer.fit_transform(texts)
    return vectorizer, matrix, rows

@st.cache_data(ttl=120, show_spinner=False)
def search_documents(query: str, category="All Categories", top_k=10):
    query = query.strip()
    if not query:
        return []
    vectorizer, matrix, rows = build_search_index()
    if vectorizer is None:
        return []

    q_vec = vectorizer.transform([query])
    scores = cosine_similarity(q_vec, matrix).flatten()

    results = []
    for idx, score in enumerate(scores):
        if score <= 0.01:
            continue
        row = rows[idx]
        if category != "All Categories" and row["category"] != category:
            continue
        results.append({
            **row,
            "raw_score": float(score),
            "snippet": make_snippet(row["text"], query),
            "direct_sentence": extract_direct_sentence(row["text"], query)
        })

    results.sort(key=lambda x: x["raw_score"], reverse=True)
    return results[:top_k]

@st.cache_data(ttl=300, show_spinner=False)
def render_pdf_page_highlighted(doc_path: str, page_num: int, target_text: str = "", zoom_level: int = 150):
    try:
        with fitz.open(doc_path) as doc:
            page_idx = max(0, min(page_num - 1, len(doc) - 1))
            page = doc[page_idx]

            if target_text:
                clean_target = " ".join(target_text.split()[:12])
                rects = page.search_for(clean_target)

                if not rects and len(target_text.split()) > 3:
                    for part in target_text.split(". "):
                        if len(part.strip()) > 10:
                            rects.extend(page.search_for(part.strip()[:40]))

                for r in rects:
                    annot = page.add_highlight_annot(r)
                    annot.set_colors(stroke=[0.0, 0.66, 0.51])
                    annot.update()

            scale = (zoom_level / 100.0) * 1.3
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            return pix.tobytes("png")
    except Exception:
        return None

def clear_all_caches():
    build_search_index.clear()
    search_documents.clear()
    get_documents.clear()
    get_categories.clear()
    render_pdf_page_highlighted.clear()

# ============================================================
# TOPBAR WITH SIDE-BY-SIDE ADMIN POPUP
# ============================================================

with st.container(key="topbar_container"):
    head_left, head_right = st.columns([3.5, 1.5], vertical_alignment="center")

    with head_left:
        st.markdown(
            """
            <div style="display: flex; align-items: center; gap: 16px;">
                <div style="width: 140px;">
                    <div class="brand-mark"></div>
                    <div style="font-size: 13px; font-weight: 700; line-height: 1.1; color: #111;">HEWLETT PACKARD<br>ENTERPRISE</div>
                </div>
                <div style="height: 40px; width: 1px; background: #b8c7cf;"></div>
                <div>
                    <div style="font-size: 24px; font-weight: 700; color: #102d42; line-height: 1.1;">Knowledge Base</div>
                    <div style="font-size: 12.5px; color: #304a5c; margin-top: 3px;">Find exact information from your organization's documents</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with head_right:
        user_col, gear_col = st.columns([2.5, 1], vertical_alignment="center")

        with user_col:
            st.markdown(
                '<div style="text-align: right;"><span class="top-user-pill">👤 Authorized User</span></div>',
                unsafe_allow_html=True,
            )

        with gear_col:
            with st.popover("⚙", use_container_width=True):
                st.markdown("**Workspace Options**")
                st.caption("Active Session: Authenticated")
                st.divider()

                if st.session_state.admin_authenticated:
                    if st.button("📁 Document Center", use_container_width=True):
                        st.session_state.page = "Manage Documents"
                        st.rerun()
                    if st.button("Lock Admin Mode", use_container_width=True):
                        st.session_state.admin_authenticated = False
                        st.session_state.page = "Search"
                        st.rerun()
                else:
                    if st.button("🔒 Admin Control Panel", use_container_width=True):
                        st.session_state.page = "Admin Login"
                        st.rerun()

                if st.button("Revoke Browser Access", use_container_width=True):
                    clear_browser_access()
                    st.session_state.admin_authenticated = False
                    st.rerun()

# ============================================================
# ADMIN LOGIN VIEW
# ============================================================

if st.session_state.page == "Admin Login":
    st.markdown("### Administrator Verification")
    st.caption("Enter the management PIN to ingest or purge organizational documents.")
    with st.form("admin_pin_form"):
        pin = st.text_input("Master PIN", type="password", placeholder="Enter PIN")
        col_sub, col_cancel = st.columns([1, 1])
        with col_sub:
            sub = st.form_submit_button("Verify", type="primary", use_container_width=True)
        with col_cancel:
            cancel = st.form_submit_button("Cancel", use_container_width=True)

    if sub:
        if pin == get_admin_pin():
            st.session_state.admin_authenticated = True
            st.session_state.page = "Manage Documents"
            st.rerun()
        else:
            st.error("Access denied: Invalid PIN.")
    if cancel:
        st.session_state.page = "Search"
        st.rerun()

# ============================================================
# DOCUMENT MANAGEMENT VIEW
# ============================================================

elif st.session_state.page == "Manage Documents":
    if not st.session_state.admin_authenticated:
        st.session_state.page = "Admin Login"
        st.rerun()

    c_head1, c_head2 = st.columns([4, 1])
    with c_head1:
        st.markdown("### Knowledge Repository Management")
        st.caption("Ingest, monitor, and remove indexed PDF sources.")
    with c_head2:
        if st.button("← Back to Search", use_container_width=True):
            st.session_state.page = "Search"
            st.rerun()

    up_col, list_col = st.columns([1, 1.4], gap="large")

    with up_col:
        st.markdown("##### Ingest Documents")
        cat = st.selectbox("Assign Category", ["General", "Policies", "Procedures", "Technical Support", "Licensing", "Product"], key="admin_cat")
        files = st.file_uploader("Upload PDF Documents", type=["pdf"], accept_multiple_files=True)
        if st.button("Index Documents", type="primary", use_container_width=True, disabled=not files):
            bar = st.progress(0)
            for i, f in enumerate(files):
                ok, msg = add_document(f.name, f.getvalue(), cat)
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)
                bar.progress((i + 1) / len(files))
            clear_all_caches()
            st.session_state.search_results = []
            st.session_state.search_signature = None

    with list_col:
        st.markdown("##### Indexed Documents")
        all_docs = get_documents()
        if not all_docs:
            st.info("No documents currently indexed.")
        else:
            for d in all_docs:
                with st.container(border=True):
                    dc1, dc2 = st.columns([4, 1])
                    with dc1:
                        st.markdown(f"**{d['filename']}**")
                        st.caption(f"{d['category']} • {d['page_count']} Pages • {format_bytes(d['file_size'])}")
                    with dc2:
                        if st.button("Delete", key=f"del_{d['id']}", use_container_width=True):
                            delete_document(d["id"])
                            clear_all_caches()
                            st.rerun()

# ============================================================
# SEARCH & SPLIT DOCUMENT INSPECTOR VIEW
# ============================================================

else:
    st.session_state.page = "Search"

    # Search Bar & Filter Controls
    sc1, sc2, sc3 = st.columns([6.2, 1.1, 1.1], gap="small")
    with sc1:
        query_val = st.text_input(
            "Search Knowledge Base",
            value=st.session_state.search_query,
            placeholder="Type exact query (e.g. 'What causes an IMC software to change its serial number?')...",
            label_visibility="collapsed",
            key="input_search_query",
        )
    with sc2:
        btn_search = st.button("Search", type="primary", use_container_width=True)
    with sc3:
        with st.popover("⚙ Filters", use_container_width=True):
            f_cat = st.selectbox("Category", ["All Categories"] + get_categories(), key="f_cat")
            f_top = st.selectbox("Max Results", [5, 10, 15], index=1, key="f_top")

    if btn_search:
        st.session_state.search_query = query_val
        st.session_state.search_signature = None
        st.session_state.force_result_id = None

    active_q = st.session_state.search_query.strip()
    active_cat = st.session_state.get("f_cat", "All Categories")
    active_top = st.session_state.get("f_top", 10)

    if active_q:
        sig = (active_q, active_cat, active_top)
        if st.session_state.search_signature != sig:
            st.session_state.search_results = search_documents(active_q, category=active_cat, top_k=active_top)
            st.session_state.search_signature = sig
            st.session_state.viewer_page = None
            st.session_state.highlight_target = ""

        results = st.session_state.search_results

        if st.session_state.force_result_id:
            forced = [r for r in results if r["id"] == st.session_state.force_result_id]
            unforced = [r for r in results if r["id"] != st.session_state.force_result_id]
            if forced:
                results = forced + unforced
                st.session_state.viewer_page = forced[0].get("page_number", 1)
                st.session_state.highlight_target = (
                    forced[0].get("direct_sentence")
                    or forced[0].get("exact_passage")
                    or extract_direct_sentence(forced[0].get("text", ""), active_q)
                )

        if not results:
            st.warning("No matching references found across the knowledge base. Try refining your keywords.")
        else:
            best_match = results[0]

            # Defensive fallback to avoid any KeyError across older sessions or caches
            direct_answer = (
                best_match.get("direct_sentence")
                or best_match.get("exact_passage")
                or extract_direct_sentence(best_match.get("text", ""), active_q)
            )

            # ----------------- HERO CARD: EXACT EXTRACTED ANSWER -----------------
            st.markdown(
                f"""
                <div class="exact-answer-card">
                    <div class="exact-answer-head">
                        <div class="exact-answer-title">Direct Extracted Answer</div>
                        <span class="match-pill">Top Confidence Match</span>
                    </div>
                    <div class="exact-answer-text">"{direct_answer}"</div>
                    <div class="answer-meta">
                        Source Document: <b>{best_match.get('filename', 'Document')}</b> &nbsp;•&nbsp; Page {best_match.get('page_number', 1)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # Two-Column Balanced Split: Left = Evidence, Right = Inspector
            left_col, right_col = st.columns([1, 1.2], gap="large")

            # ======================== LEFT: EVIDENCE RESULTS ========================
            with left_col:
                st.markdown(f"**Supporting Citations** ({len(results)} matches located)")

                for i, res in enumerate(results):
                    is_top = i == 0
                    label_text = "Top Match" if is_top else "Supporting Reference"
                    label_color = "#00a982" if is_top else "#0369a1"

                    with st.container():
                        st.markdown(
                            f"""
                            <div class="result-card">
                                <span class="result-label" style="color:{label_color};">{label_text}</span>
                                <div class="result-filename">{res.get('filename', 'Document')}</div>
                                <div style="font-size:11px;color:#64748b;margin-top:2px;">Page {res.get('page_number', 1)}</div>
                                <div class="result-snippet">{res.get('snippet', '')}</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        btn_c1, btn_c2 = st.columns([1.2, 2.5])
                        with btn_c1:
                            if st.button("Inspect Page", key=f"btn_res_{res.get('id', i)}_{i}", use_container_width=True):
                                st.session_state.viewer_page = res.get("page_number", 1)
                                st.session_state.highlight_target = (
                                    res.get("direct_sentence")
                                    or res.get("exact_passage")
                                    or extract_direct_sentence(res.get("text", ""), active_q)
                                )
                                st.rerun()
                        st.write("")

            # ======================== RIGHT: PDF VIEWER ========================
            with right_col:
                current_page = st.session_state.viewer_page or best_match.get("page_number", 1)
                highlight_phrase = st.session_state.highlight_target or direct_answer
                doc_path = best_match.get("stored_path")

                # Get page boundaries safely
                max_pages = 1
                if doc_path and os.path.exists(doc_path):
                    with fitz.open(doc_path) as d_meta:
                        max_pages = len(d_meta)

                current_page = max(1, min(current_page, max_pages))

                # Header Navigation Toolbar
                st.markdown(
                    f"""
                    <div class="reader-toolbar">
                        <div class="reader-title">📄 {best_match.get('filename', 'Document')}</div>
                        <div class="reader-meta">Page {current_page} of {max_pages} &nbsp;•&nbsp; Active Highlight Anchor Active</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                t_nav1, t_nav2, t_nav3, t_nav4 = st.columns([1, 1.2, 1, 1.2])
                with t_nav1:
                    if st.button("◀ Prev", disabled=(current_page <= 1), use_container_width=True):
                        st.session_state.viewer_page = current_page - 1
                        st.session_state.highlight_target = ""
                        st.rerun()
                with t_nav2:
                    st.markdown(f"<div style='text-align:center;padding-top:7px;font-size:12px;color:#475569;'>Page <b>{current_page}</b> of {max_pages}</div>", unsafe_allow_html=True)
                with t_nav3:
                    if st.button("Next ▶", disabled=(current_page >= max_pages), use_container_width=True):
                        st.session_state.viewer_page = current_page + 1
                        st.session_state.highlight_target = ""
                        st.rerun()
                with t_nav4:
                    zoom_val = st.selectbox("Zoom", [125, 150, 175, 200], index=1, format_func=lambda x: f"{x}%", label_visibility="collapsed")

                # High-DPI Highlight Rendering
                if doc_path and os.path.exists(doc_path):
                    img_data = render_pdf_page_highlighted(
                        doc_path=doc_path,
                        page_num=current_page,
                        target_text=highlight_phrase,
                        zoom_level=zoom_val,
                    )
                    if img_data:
                        st.image(img_data, use_container_width=True)
                        st.caption("🟢 Exact matching sentence automatically targeted in emerald highlight.")
                    else:
                        st.error("Failed to render PDF page canvas.")
                else:
                    st.error("Document file not found on disk.")
    else:
        docs = get_documents()
        tot_pages = sum(d["page_count"] for d in docs)
        st.markdown(
            f"""
            <div class="welcome-card">
                <div class="welcome-icon">⌕</div>
                <div class="welcome-title">Search Organization Knowledge Base</div>
                <div class="welcome-text">Type a procedure name, error code, hardware SKU, or compliance policy to pull exact verbatim text and inspect the original source PDF page.</div>
                <div class="welcome-stats">
                    <span><b>{len(docs)}</b> Documents Indexed</span>
                    <span><b>{tot_pages:,}</b> Searchable Pages</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
