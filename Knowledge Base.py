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

try:
    from streamlit_cookies_controller import CookieController
except Exception:
    CookieController = None
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# Optional AI support:
# pip install openai
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

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
:root{--navy:#082536;--teal:#00a982;--teal2:#008c76;--bg:#f4f7f9;--line:#dfe7eb;--text:#0b2438;--muted:#6b7e8b}
html,body,[class*="css"]{font-family:Arial,Helvetica,sans-serif}.stApp{background:var(--bg);color:var(--text)}
[data-testid="stHeader"]{background:transparent;height:0;min-height:0}[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stStatusWidget"],[data-testid="stAppDeployButton"],[data-testid="stMainMenu"],[data-testid="stHeaderActionElements"]{display:none!important}
[data-testid="stSidebar"]{display:block!important;background:linear-gradient(180deg,#082b35,#06252e);border-right:1px solid #0c4a55}[data-testid="stSidebar"]>div{background:transparent}[data-testid="stSidebarContent"]{padding:18px 12px 20px}
.sidebar-logo{font-size:36px;font-weight:900;letter-spacing:-3px;color:#fff;line-height:1;margin:0 0 4px 8px}.sidebar-logo span{color:#00a982}.sidebar-kicker{color:#a9c6ce;font-size:10px;margin-left:9px;margin-bottom:22px}.sidebar-section{font-size:10px;text-transform:uppercase;letter-spacing:.12em;color:#6f9ba4;margin:18px 8px 5px}.sidebar-footer{position:fixed;bottom:18px;color:#cfe1e6;margin-left:8px;font-size:12px}.sidebar-footer b{font-size:15px;color:white}
[data-testid="stSidebar"] button{border:0!important;background:transparent!important;color:#f2f8fa!important;text-align:left!important;border-radius:8px!important;min-height:42px!important;font-size:14px!important;padding:7px 10px!important;margin:2px 0!important}[data-testid="stSidebar"] button:hover{background:rgba(0,169,130,.18)!important}.sidebar-active button{background:linear-gradient(90deg,#00a982,#009777)!important;color:#fff!important}
.top-title{font-size:27px;font-weight:800;color:#0a2034;padding-top:6px}.top-user-wrap{display:flex;justify-content:flex-end;align-items:center;gap:8px}.avatar{width:38px;height:38px;border-radius:50%;background:linear-gradient(135deg,#efc19e,#cf8b66);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:12px}.user-name{font-size:12px;font-weight:700;color:#0b2538}.user-role{font-size:11px;color:#6d7f8b}
.st-key-top_user_bar{display:flex!important;align-items:center!important;justify-content:flex-end!important;background:#eaf7f5!important;border-radius:6px!important;padding:4px 6px 4px 10px!important;min-height:38px!important;width:max-content!important;max-width:100%!important;margin-left:auto!important}.st-key-top_user_bar [data-testid="column"]{padding:0!important;margin:0!important}.st-key-top_user_bar .top-user-wrap{display:flex;align-items:center;justify-content:flex-end;min-height:30px}.st-key-top_user_bar .user-name{font-size:12px;white-space:nowrap;color:#315468;font-weight:600}.st-key-top_user_bar .admin-gear-anchor{display:block;width:1px;height:1px}.st-key-top_user_bar [data-testid="stPopover"]{display:flex;align-items:center;justify-content:center}.st-key-top_user_bar [data-testid="stPopover"] > button{margin:0!important;min-height:30px!important;height:30px!important;width:30px!important;padding:0!important;border:0!important;border-radius:6px!important;background:#eaf7f5!important;color:#0b5260!important;box-shadow:none!important;font-size:15px!important;line-height:1!important}.st-key-top_user_bar [data-testid="stPopover"] > button:hover,.st-key-top_user_bar [data-testid="stPopover"] > button:focus{background:#d8efeb!important;color:#087c63!important;box-shadow:none!important}
.hero{background:linear-gradient(115deg,#082b35 0%,#0b5560 48%,#0c7778 100%);border-radius:10px;padding:28px 36px 20px;color:#fff;position:relative;overflow:hidden;min-height:185px;box-shadow:0 5px 20px rgba(0,40,50,.10)}.hero:after{content:"";position:absolute;right:-20px;bottom:-80px;width:470px;height:240px;background:linear-gradient(160deg,transparent 20%,rgba(0,206,190,.35) 21%,transparent 23%,rgba(0,206,190,.2) 40%,transparent 42%),linear-gradient(90deg,transparent 35%,rgba(0,206,190,.22) 36%,transparent 38%);transform:skewX(-20deg)}.hero h1{font-size:36px;line-height:1.05;margin:0 0 6px;font-weight:800;position:relative;z-index:1}.hero p{font-size:16px;margin:0;color:#e4f5f6;position:relative;z-index:1}.popular{margin:9px 0 0;font-size:11px;color:#e2f2f3}.chip{display:inline-block;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.12);padding:6px 11px;border-radius:18px;margin:4px 4px 0 0;color:#fff}
.metric-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:16px 0}.metric{border:1px solid var(--line);border-radius:11px;padding:15px 18px;background:#fff;min-height:84px;display:flex;align-items:center;justify-content:space-between}.metric.green{background:linear-gradient(110deg,#e8faf5,#fff)}.metric.blue{background:linear-gradient(110deg,#eaf4ff,#fff)}.metric.gold{background:linear-gradient(110deg,#fff7df,#fff)}.metric.purple{background:linear-gradient(110deg,#f4efff,#fff)}.metric-icon{width:48px;height:48px;border-radius:15px;display:flex;align-items:center;justify-content:center;font-size:23px;background:#d7f5ec}.metric-value{font-size:25px;font-weight:800;color:#0a2438}.metric-label{font-size:12px;color:#526877}.metric-arrow{font-size:22px;color:#132f42}.blue .metric-icon{background:#dcecff}.gold .metric-icon{background:#ffebbb}.purple .metric-icon{background:#e9ddff}
.section-card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:14px;box-shadow:0 2px 8px rgba(10,40,55,.03)}.section-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}.section-title{font-size:17px;font-weight:800;color:#0b2438}.section-link{font-size:12px;color:#0b2438}.cat-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.cat{border:1px solid #e5ebef;border-radius:10px;padding:10px 12px;display:flex;align-items:center;gap:10px;min-height:66px}.cat-icon{width:40px;height:40px;border-radius:50%;background:#e6f8f2;display:flex;align-items:center;justify-content:center;font-size:19px;color:#008c76}.cat-name{font-size:12px;font-weight:700;color:#122b3d}.cat-count{font-size:10px;color:#6d7f8b;margin-top:3px}.recent-row{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid #edf1f3}.recent-row:last-child{border-bottom:0}.pdf-icon{width:34px;height:34px;border-radius:9px;background:#fff0f0;color:#e5483f;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800}.recent-title{font-size:12px;font-weight:700;color:#0e2a3c}.recent-meta,.recent-date{font-size:10px;color:#70818d}.recent-date{margin-left:auto;white-space:nowrap}.featured-title{font-size:16px;font-weight:800;color:#0b2438;margin-top:8px}.featured-meta{font-size:11px;color:#6c7e8a;margin-top:3px}
.search-count{color:#687b87;font-size:11px;margin:5px 0 8px}.reader-toolbar{background:#fff;border:1px solid var(--line);border-radius:8px;padding:8px 10px;margin-bottom:8px}.reader-title{font-size:15px;font-weight:700}.reader-meta{font-size:11px;color:var(--muted);margin-top:2px}.reader-match{background:#e7f8f1;border:1px solid #9bdcc8;color:#087c63;border-radius:5px;padding:5px 8px;font-size:10px;font-weight:700;display:inline-block;margin-top:5px}.match-panel{background:#fff;border:1px solid var(--line);border-radius:8px;padding:10px;margin-top:10px}.match-panel-title{font-size:13px;font-weight:700;margin-bottom:7px}.match-item{background:#f7fafb;border:1px solid #e1e8ec;border-radius:6px;padding:7px 8px;margin-bottom:6px}.match-item-page{color:#0561a0;font-size:10px;font-weight:700}.match-item-text{color:#385362;font-size:10px;line-height:1.35;margin-top:2px}.result-snippet{color:#536b78;font-size:10px;line-height:1.4;margin-top:5px;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}.result-page{color:#687b87;font-size:10px;margin-top:3px}.result-score-pill{float:right;background:#e7f8f1;color:#087c63;border-radius:10px;padding:2px 6px;font-size:9px;font-weight:700}.panel-title{font-size:17px;font-weight:800;color:#0b2438;margin:3px 0 8px}
[data-testid="stVerticalBlockBorderWrapper"]{border-color:var(--line)!important;border-radius:9px!important}.st-key-search_result_best [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_1 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_2 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_3 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_4 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_5 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_6 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_7 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_8 [data-testid="stVerticalBlockBorderWrapper"],.st-key-search_result_9 [data-testid="stVerticalBlockBorderWrapper"]{padding:8px 10px!important}.result-label{color:#0561a0;font-weight:700;font-size:11px}.result-filename{margin-top:3px;color:#0561a0;font-weight:700;font-size:12px;line-height:1.35;word-break:break-word}.st-key-search_result_best button,.st-key-search_result_1 button,.st-key-search_result_2 button,.st-key-search_result_3 button,.st-key-search_result_4 button,.st-key-search_result_5 button,.st-key-search_result_6 button,.st-key-search_result_7 button,.st-key-search_result_8 button,.st-key-search_result_9 button{min-height:28px!important;height:28px!important;padding:2px 8px!important;font-size:11px!important;margin-top:4px!important}
.page-heading{display:flex;justify-content:space-between;align-items:center;margin:18px 0}.page-title{font-size:26px;font-weight:700}.page-description{color:var(--muted);font-size:13px;margin-top:4px}.admin-badge{background:#e4f6f0;color:#087c63;font-size:11px;font-weight:700;border-radius:20px;padding:6px 12px}.admin-card{max-width:420px;margin:80px auto 20px;text-align:center}.admin-icon{font-size:40px;color:var(--teal)}.admin-title{font-size:24px;font-weight:700}.admin-subtitle{color:var(--muted);margin-top:5px;font-size:13px}
[data-testid="stFileUploader"]{background:#fff;border-radius:10px;border:1px dashed #9ab1bc}button[kind="primary"]{background:var(--teal)!important;border-color:var(--teal)!important}button[kind="primary"]:hover{background:var(--teal2)!important}.stDownloadButton button{border-color:#00a982!important;color:#087b64!important}
@media(max-width:1100px){.metric-grid,.cat-grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:700px){.metric-grid,.cat-grid{grid-template-columns:1fr}.hero{padding:22px}.hero h1{font-size:26px}.top-title{font-size:21px}}
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

    results.sort(key=lambda x: x["score"], reverse=True)

    return results[:top_k]


# ============================================================
# AI Q&A
# ============================================================

def get_openai_client():
    api_key = os.getenv("OPENAI_API_KEY")

    if not api_key or OpenAI is None:
        return None

    try:
        return OpenAI(api_key=api_key)
    except Exception:
        return None


def extractive_answer(question, results):
    if not results:
        return (
            "I could not find relevant information in the uploaded knowledge base."
        )

    selected = results[:5]

    answer_parts = []

    for item in selected:
        text = item["text"].strip()

        # Keep answer reasonably concise.
        if len(text) > 700:
            text = text[:700].rsplit(" ", 1)[0] + "..."

        answer_parts.append(text)

    return "\n\n".join(answer_parts)


def ai_answer(question, results):
    if not results:
        return (
            "I could not find relevant information in the uploaded documents.",
            [],
        )

    client = get_openai_client()

    if client is None:
        return extractive_answer(question, results), results[:5]

    context_blocks = []

    for i, item in enumerate(results[:8], start=1):
        context_blocks.append(
            f"""
SOURCE {i}
Document: {item['filename']}
Page: {item['page_number']}
Category: {item['category']}

CONTENT:
{item['text']}
"""
        )

    context = "\n".join(context_blocks)

    system_prompt = """
You are a company knowledge-base assistant.

Answer the user's question using ONLY the provided document context.
Do not invent policies, procedures, facts, dates, or instructions.
If the documents do not contain enough information, explicitly say that
the knowledge base does not provide enough information.

Keep answers concise and practical.
When appropriate, use numbered steps or bullet points.

Do not cite a source that was not provided in the context.
"""

    user_prompt = f"""
QUESTION:
{question}

DOCUMENT CONTEXT:
{context}

Answer the question using only the document context.
"""

    try:
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.1,
        )

        answer = response.choices[0].message.content.strip()
        return answer, results[:8]

    except Exception as e:
        # Graceful fallback if AI service is unavailable.
        return (
            "AI answering is temporarily unavailable. "
            "Here are the most relevant passages from the knowledge base:\n\n"
            + extractive_answer(question, results)
        ), results[:5]


# ============================================================
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

if "force_result_id" not in st.session_state:
    st.session_state.force_result_id = None

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


# ============================================================
# APP SHELL
# ============================================================

def go(page_name):
    st.session_state.page = page_name
    if page_name != "Search":
        st.session_state.search_query = ""
    st.rerun()

with st.sidebar:
    st.markdown('<div class="sidebar-logo">HP<span>E</span></div><div class="sidebar-kicker">KNOWLEDGE BASE</div>', unsafe_allow_html=True)
    nav_items=[("Home","⌂","Home"),("Browse All","▤","Browse All"),("Categories","▦","Categories"),("Favorites","☆","Favorites"),("Recent","◷","Recent")]
    for label,icon,target in nav_items:
        active=(st.session_state.page==target) or (target=="Home" and st.session_state.page=="Search" and not st.session_state.search_query)
        if active: st.markdown('<div class="sidebar-active">',unsafe_allow_html=True)
        if st.button(f"{icon}   {label}",key=f"nav_{target}",use_container_width=True): go(target)
        if active: st.markdown('</div>',unsafe_allow_html=True)
    st.markdown('<div class="sidebar-section">Workspace</div>',unsafe_allow_html=True)
    for label,icon,target in [("Upload PDF","⇧","Upload PDF"),("Manage Content","▤","Manage Documents"),("Analytics","⌁","Analytics")]:
        if st.button(f"{icon}   {label}",key=f"nav_{target}",use_container_width=True):
            if target in ("Upload PDF","Manage Documents") and not st.session_state.admin_authenticated:
                st.session_state.page="Admin Login"; st.rerun()
            else: go(target)
    st.markdown('<div class="sidebar-section">Support</div>',unsafe_allow_html=True)
    for label,icon,target in [("Feedback","▢","Feedback"),("Help","?","Help")]:
        if st.button(f"{icon}   {label}",key=f"nav_{target}",use_container_width=True): go(target)
    st.markdown('<div class="sidebar-footer"><b>HPE</b><br>Knowledge Base<br><span style="color:#6f9ba4">v1.0.0</span></div>',unsafe_allow_html=True)

left,search,user=st.columns([3.0,4.7,2.1],gap="small")
with left:
    st.markdown('<div class="top-title">Knowledge Base</div>',unsafe_allow_html=True)
with search:
    top_query=st.text_input("Header search",value=st.session_state.search_query,placeholder="Search for topics, keywords, or questions...",label_visibility="collapsed",key="header_search_box")
with user:
    # Keep the Authorized User label and admin gear in one compact Streamlit
    # container so the gear is physically beside the text and cannot affect
    # the main page layout or the search Filters popover.
    with st.container(key="top_user_bar"):
        user_name_col, gear_col = st.columns([1, 0.28], gap="small", vertical_alignment="center")
        with user_name_col:
            st.markdown('<div class="top-user-wrap"><div class="user-name">Authorized User</div></div>',unsafe_allow_html=True)
        with gear_col:
            with st.popover("⚙",use_container_width=False):
                st.markdown("**Knowledge Base Access**")
                st.caption("Browser authorization: persistent until manually cleared")
                if st.session_state.admin_authenticated:
                    if st.button("Manage Documents",use_container_width=True,key="gear_manage"): go("Manage Documents")
                    if st.button("Sign out admin",use_container_width=True,key="gear_signout"):
                        st.session_state.admin_authenticated=False; go("Search")
                else:
                    if st.button("🔒 Manage Documents",use_container_width=True,key="gear_login"): go("Admin Login")
                if st.button("Clear Browser Access",use_container_width=True,key="gear_clear"):
                    clear_browser_access(); st.session_state.admin_authenticated=False; st.rerun()
if top_query.strip()!=st.session_state.search_query.strip():
    st.session_state.search_query=top_query.strip(); st.session_state.page="Search"; st.session_state.search_signature=None
    if top_query.strip(): st.rerun()
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
                            st.rerun()

    if st.button("← Back to Search"):
        st.session_state.page = "Search"
        st.rerun()


# ============================================================
# HOME / SEARCH
# ============================================================
if st.session_state.page in {"Browse All","Categories","Favorites","Recent","Analytics","Feedback","Help","Upload PDF"}:
    docs=get_documents()
    if st.session_state.page=="Upload PDF":
        if not st.session_state.admin_authenticated: st.session_state.page="Admin Login"; st.rerun()
        else: st.session_state.page="Manage Documents"; st.rerun()
    elif st.session_state.page=="Browse All":
        st.markdown('<div class="page-heading"><div><div class="page-title">Browse All Documents</div><div class="page-description">Search and open any indexed PDF.</div></div></div>',unsafe_allow_html=True)
        q=st.text_input("Browse",placeholder="Filter documents by title...",label_visibility="collapsed",key="browse_filter")
        for doc in docs:
            if q and q.lower() not in doc["filename"].lower(): continue
            a,b=st.columns([5,1])
            with a: st.markdown(f"**📄 {doc['filename']}**"); st.caption(f"{doc['category']} · {doc['page_count']} pages · {format_bytes(doc['file_size'])}")
            with b:
                if st.button("Open",key=f"browse_{doc['id']}"):
                    st.session_state.search_query=doc["filename"].replace(".pdf",""); st.session_state.force_result_id=None; st.session_state.page="Search"; st.session_state.search_signature=None; st.rerun()
                if st.button("☆",key=f"fav_{doc['id']}"):
                    favs=st.session_state.setdefault("favorites",set())
                    if doc["id"] in favs: favs.remove(doc["id"])
                    else: favs.add(doc["id"])
                    st.rerun()
            st.divider()
    elif st.session_state.page=="Categories":
        st.markdown('<div class="page-heading"><div><div class="page-title">Categories</div><div class="page-description">Browse documents by category.</div></div></div>',unsafe_allow_html=True)
        counts={}
        for d in docs: counts[d["category"]]=counts.get(d["category"],0)+1
        cols=st.columns(4)
        for i,(cat,count) in enumerate(sorted(counts.items())):
            with cols[i%4]:
                with st.container(border=True):
                    st.markdown(f"### {cat}"); st.caption(f"{count} document(s)")
                    if st.button("Browse",key=f"cat_{cat}"):
                        st.session_state.search_query=cat; st.session_state.page="Search"; st.session_state.search_signature=None; st.rerun()
    elif st.session_state.page=="Recent":
        st.markdown('<div class="page-heading"><div><div class="page-title">Recent Documents</div><div class="page-description">Latest documents added to the knowledge base.</div></div></div>',unsafe_allow_html=True)
        for d in docs[:20]: st.markdown(f"**📄 {d['filename']}** — {d['category']} · {d['page_count']} pages"); st.divider()
    elif st.session_state.page=="Favorites":
        st.markdown('<div class="page-heading"><div><div class="page-title">Favorites</div><div class="page-description">Documents marked for quick access in this browser session.</div></div></div>',unsafe_allow_html=True)
        favs=st.session_state.get("favorites",set()); favorite_docs=[d for d in docs if d["id"] in favs]
        if not favorite_docs: st.info("No favorites yet.")
        for d in favorite_docs: st.markdown(f"⭐ **{d['filename']}**")
    elif st.session_state.page=="Analytics":
        st.markdown('<div class="page-heading"><div><div class="page-title">Analytics</div><div class="page-description">Current knowledge-base inventory.</div></div></div>',unsafe_allow_html=True)
        a,b,c,d=st.columns(4); a.metric("Documents",len(docs)); b.metric("Categories",len(get_categories())); c.metric("Pages",sum(int(x["page_count"] or 0) for x in docs)); d.metric("Storage",format_bytes(sum(int(x["file_size"] or 0) for x in docs)))
    elif st.session_state.page=="Feedback":
        st.markdown('<div class="page-heading"><div><div class="page-title">Feedback</div><div class="page-description">Tell us what would make search easier.</div></div></div>',unsafe_allow_html=True)
        with st.form("feedback_form"):
            feedback=st.text_area("Feedback",placeholder="What should we improve?")
            if st.form_submit_button("Submit Feedback",type="primary"): st.success("Thank you for your feedback.")
    elif st.session_state.page=="Help":
        st.markdown('<div class="page-heading"><div><div class="page-title">Help</div><div class="page-description">Quick guide to using the knowledge base.</div></div></div>',unsafe_allow_html=True)
        st.markdown("**Search:** enter a topic, keyword, phrase, or question.  \n**Open Result:** opens the matching PDF page with highlighted terms.  \n**Filters:** narrow results by category and result count.  \n**Admin:** use the gear beside Authorized User, then enter the Admin PIN to manage PDFs.")
else:
    docs=get_documents(); cats=get_categories(); active_query=st.session_state.search_query.strip()
    if not active_query:
        st.markdown('<div class="hero"><h1>Find the answers you need</h1><p>Search our knowledge base, explore topics, or browse by category.</p></div>',unsafe_allow_html=True)
        hq,hb=st.columns([5.4,1],gap="small")
        with hq: hero_query=st.text_input("Hero search",placeholder="Search for solutions, guides, or keywords...",label_visibility="collapsed",key="hero_search_box")
        with hb:
            if st.button("Search",type="primary",use_container_width=True,key="hero_search_btn"):
                st.session_state.search_query=hero_query.strip(); st.session_state.page="Search"; st.session_state.search_signature=None; st.rerun()
        st.markdown('<div class="popular">Popular searches: <span class="chip">Licensing</span><span class="chip">Portal Access</span><span class="chip">Account Setup</span><span class="chip">Troubleshooting</span><span class="chip">HPE GreenLake</span><span class="chip">Software Support</span></div>',unsafe_allow_html=True)
        total_pages=sum(int(d["page_count"] or 0) for d in docs); recent_count=min(24,len(docs))
        st.markdown(f'<div class="metric-grid"><div class="metric green"><div class="metric-icon">▤</div><div><div class="metric-value">{len(docs)}</div><div class="metric-label">Total Documents</div></div><div class="metric-arrow">›</div></div><div class="metric blue"><div class="metric-icon">▱</div><div><div class="metric-value">{len(cats)}</div><div class="metric-label">Categories</div></div><div class="metric-arrow">›</div></div><div class="metric gold"><div class="metric-icon">★</div><div><div class="metric-value">{total_pages:,}</div><div class="metric-label">Indexed Pages</div></div><div class="metric-arrow">›</div></div><div class="metric purple"><div class="metric-icon">⇧</div><div><div class="metric-value">{recent_count}</div><div class="metric-label">Recently Added</div></div><div class="metric-arrow">›</div></div></div>',unsafe_allow_html=True)
        counts={}
        for d in docs: counts[d["category"]]=counts.get(d["category"],0)+1
        items=sorted(counts.items(),key=lambda x:(-x[1],x[0]))[:8]
        cat_html=''.join([f'<div class="cat"><div class="cat-icon">{["●","⌕","▣","⚙","☁","▤","◈","◇"][i]}</div><div><div class="cat-name">{cat}</div><div class="cat-count">{count} documents</div></div></div>' for i,(cat,count) in enumerate(items)])
        st.markdown(f'<div class="section-card"><div class="section-head"><div class="section-title">▦ &nbsp;Browse by Category</div><div class="section-link">{len(cats)} categories</div></div><div class="cat-grid">{cat_html}</div></div>',unsafe_allow_html=True)
        recent=sorted(docs,key=lambda x:x.get("uploaded_at","") or "",reverse=True)[:5]
        left,right=st.columns([1.05,.95],gap="large")
        with left:
            st.markdown('<div class="section-card"><div class="section-head"><div class="section-title">▤ &nbsp;Recent Documents</div><div class="section-link">View All →</div></div>',unsafe_allow_html=True)
            if recent:
                for d in recent: st.markdown(f'<div class="recent-row"><div class="pdf-icon">PDF</div><div><div class="recent-title">{d["filename"]}</div><div class="recent-meta">{d["category"]} · {d["page_count"]} pages</div></div><div class="recent-date">{str(d.get("uploaded_at", ""))[:10]}</div></div>',unsafe_allow_html=True)
            else: st.caption("No documents have been indexed yet.")
            st.markdown('</div>',unsafe_allow_html=True)
        with right:
            st.markdown('<div class="section-card"><div class="section-head"><div class="section-title">★ &nbsp;Featured Document</div></div>',unsafe_allow_html=True)
            if docs:
                featured=recent[0] if recent else docs[0]; preview=render_pdf_page(featured["stored_path"],1,scale=1.0)
                if preview: st.image(preview,use_container_width=True)
                st.markdown(f'<div class="featured-title">{featured["filename"]}</div><div class="featured-meta">{featured["category"]} · {featured["page_count"]} pages</div>',unsafe_allow_html=True)
                if st.button("Open Document →",type="primary",use_container_width=True,key="featured_open"):
                    st.session_state.search_query=featured["filename"].replace(".pdf",""); st.session_state.page="Search"; st.session_state.search_signature=None; st.rerun()
            else: st.info("Upload a PDF to feature it here.")
            st.markdown('</div>',unsafe_allow_html=True)
    else:
        search_col,button_col,filter_col=st.columns([6.4,1,1],gap="small")
        with search_col: query=st.text_input("Search",value=active_query,placeholder="Search for solutions, guides, or keywords...",label_visibility="collapsed",key="main_search_box")
        with button_col: search_clicked=st.button("Search",type="primary",use_container_width=True,key="result_search_btn")
        with filter_col:
            with st.popover("☷ Filters",use_container_width=True):
                category=st.selectbox("Category",["All Categories"]+get_categories(),key="search_category"); top_k=st.selectbox("Results",[5,10,20],index=1,key="search_top_k")
        if search_clicked: st.session_state.search_query=query.strip(); st.session_state.search_signature=None; st.rerun()
        active_query=st.session_state.search_query.strip(); category=st.session_state.get("search_category","All Categories"); top_k=st.session_state.get("search_top_k",10); sig=(active_query,category,top_k)
        if st.session_state.get("search_signature")!=sig: st.session_state.search_results=search_documents(active_query,category=category,top_k=top_k); st.session_state.search_signature=sig; st.session_state.viewer_page=None
        results=st.session_state.get("search_results",[])
        if st.session_state.force_result_id is not None:
            fid=st.session_state.force_result_id; forced=[r for r in results if r["id"]==fid]; others=[r for r in results if r["id"]!=fid]
            if forced: results=forced+others; st.session_state.viewer_page=forced[0]["page_number"]
            st.session_state.force_result_id=None
        if not results: st.warning("No matching PDF was found. Try different keywords or upload another document.")
        else:
            best=results[0]; best_doc=get_document_by_id(best["document_id"]); total_pages=int(best_doc["page_count"] or 0) if best_doc else 0
            if st.session_state.get("viewer_page") is None: st.session_state.viewer_page=int(best["page_number"])
            viewer_page=max(1,min(int(st.session_state.viewer_page),max(total_pages,1)))
            st.markdown(f'<div class="search-count">{len(results)} search result(s) · Highest match shown first</div>',unsafe_allow_html=True)
            search_results_col, source_col = st.columns([0.88, 1.92], gap="large")
            with search_results_col:
                # Search Results intentionally stay on the LEFT; PDF reader is on the RIGHT.
                st.markdown('<div class="panel-title">Search Results</div>',unsafe_allow_html=True)
                for i,result in enumerate(results):
                    with st.container(key=f"search_result_{'best' if i==0 else i}",border=True):
                        score_pct=min(99,max(1,round(result["score"]*100))); label="Highest Match" if i==0 else "Search Result"
                        st.markdown(f'<div class="result-label">{label}<span class="result-score-pill">{score_pct}%</span></div><div class="result-filename">{result["filename"]}</div><div class="result-page">Page {result["page_number"]}</div><div class="result-snippet">{result["snippet"]}</div>',unsafe_allow_html=True)
                        if st.button("Open Result",key=f"open_result_{result['id']}",use_container_width=True): st.session_state.force_result_id=result["id"]; st.session_state.viewer_page=result["page_number"]; st.rerun()
                matches=find_document_matches(best["stored_path"],active_query,limit=8)
                st.markdown('<div class="match-panel"><div class="match-panel-title">Matches in this PDF</div>',unsafe_allow_html=True)
                if matches:
                    for match in matches:
                        if st.button(f"Page {match['page']}",key=f"jump_match_{best['id']}_{match['page']}",use_container_width=True): st.session_state.viewer_page=match["page"]; st.rerun()
                        st.markdown(f'<div class="match-item"><div class="match-item-page">Page {match["page"]} · {match["score"]} term match(es)</div><div class="match-item-text">{match["snippet"]}</div></div>',unsafe_allow_html=True)
                else: st.caption("No exact text occurrence was detected on the other pages.")
                st.markdown('</div>',unsafe_allow_html=True)
            with source_col:
                st.markdown(f'<div class="reader-toolbar"><div class="reader-title">▣ {best["filename"]}</div><div class="reader-meta">Page {viewer_page} of {total_pages} · Search: “{active_query}”</div><div class="reader-match">Highest match · source text shown exactly as it appears in the PDF</div></div>',unsafe_allow_html=True)
                n1,n2,n3,n4=st.columns([1,1.1,1,1.1],gap="small")
                with n1:
                    if st.button("‹ Previous",disabled=viewer_page<=1,use_container_width=True,key="pdf_prev"): st.session_state.viewer_page=max(1,viewer_page-1); st.rerun()
                with n2: st.markdown(f"<div style='text-align:center;padding-top:8px;font-size:11px;color:#687b87'>Page <b>{viewer_page}</b> / {total_pages}</div>",unsafe_allow_html=True)
                with n3:
                    if st.button("Next ›",disabled=viewer_page>=total_pages,use_container_width=True,key="pdf_next"): st.session_state.viewer_page=min(total_pages,viewer_page+1); st.rerun()
                with n4: zoom=st.selectbox("Zoom",[100,125,150,175,200],index=1,format_func=lambda x:f"{x}%",label_visibility="collapsed",key="pdf_zoom")
                image=render_pdf_page_highlighted(best["stored_path"],viewer_page,active_query,scale=zoom/100*1.25)
                if image: st.image(image,width={100:720,125:820,150:980,175:1140,200:1300}.get(int(zoom),820))
                else: st.error("Unable to render this PDF page.")
                try:
                    with fitz.open(best["stored_path"]) as pdf: current_text=clean_text(pdf[viewer_page-1].get_text("text"))
                except Exception: current_text=""
                if current_text: st.markdown(f'<div class="match-panel"><div class="match-panel-title">Match context · Page {viewer_page}</div><div class="result-snippet" style="font-size:12px;line-height:1.55;color:#294a5c">{make_snippet(current_text,active_query,radius=520)}</div></div>',unsafe_allow_html=True)

st.markdown('<div style="height:22px"></div><div style="text-align:center;color:#78909c;font-size:10px">HPE Knowledge Base · PDF Search</div>',unsafe_allow_html=True)
