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
/* Admin gear: fixed beside Authorized User without taking layout space. */
.kb-topbar { position: relative; }
.st-key-gear_wrap {
    position:fixed !important;
    top:92px !important;
    right:24px !important;
    width:30px !important;
    height:30px !important;
    margin:0 !important;
    padding:0 !important;
    z-index:10000 !important;
    pointer-events:none !important;
}
.st-key-gear_wrap > div,
.st-key-gear_wrap div[data-testid="stPopover"] {
    width:30px !important;
    margin:0 !important;
    padding:0 !important;
    pointer-events:auto !important;
}
.st-key-gear_wrap div[data-testid="stPopover"] > button {
    width:30px !important;
    min-width:30px !important;
    height:30px !important;
    min-height:30px !important;
    padding:0 !important;
    border:0 !important;
    border-radius:4px !important;
    background:#e8f7f7 !important;
    box-shadow:none !important;
    color:#315468 !important;
    font-size:14px !important;
    line-height:30px !important;
}
.st-key-gear_wrap div[data-testid="stPopover"] > button:hover,
.st-key-gear_wrap div[data-testid="stPopover"] > button:focus {
    background:#e8f7f7 !important;
    color:#087c63 !important;
    box-shadow:none !important;
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
.bottom-nav-spacer { height: 46px; }
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
    margin-right:42px;
    white-space:nowrap;
    background:#e8f7f7;
    padding:7px 10px;
    border-radius:4px;
    position:relative;
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
.search-count { color:#687b87; font-size:11px; margin:2px 0 8px; }
.reader-toolbar {
    background:#ffffff; border:1px solid var(--border); border-radius:8px;
    padding:7px 10px; margin-bottom:8px;
}
.reader-title { font-size:15px; font-weight:700; color:var(--text); line-height:1.3; word-break:break-word; }
.reader-meta { font-size:11px; color:var(--muted); margin-top:2px; }
.reader-match {
    background:#e7f8f1; border:1px solid #9bdcc8; color:#087c63;
    border-radius:5px; padding:5px 8px; font-size:10px; font-weight:700;
    display:inline-block; margin-top:5px;
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

        CREATE TABLE IF NOT EXISTS decision_trees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            category TEXT DEFAULT 'General',
            description TEXT DEFAULT '',
            start_node TEXT NOT NULL,
            nodes_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
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
# DECISION TREE ENGINE
# ============================================================

def get_decision_trees():
    conn = db()
    rows = conn.execute("SELECT * FROM decision_trees ORDER BY name").fetchall()
    conn.close()
    return [dict(row) for row in rows]


def get_decision_tree(tree_id):
    conn = db()
    row = conn.execute("SELECT * FROM decision_trees WHERE id = ?", (tree_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def save_decision_tree(name, category, description, start_node, nodes):
    name = name.strip(); category = category.strip() or "General"; description = description.strip(); start_node = start_node.strip()
    if not name or not start_node or not isinstance(nodes, dict) or start_node not in nodes:
        return False, "Tree name, start node, and a valid start node definition are required."
    now = datetime.now().isoformat(timespec="seconds")
    payload = json.dumps(nodes, ensure_ascii=False, indent=2)
    conn = db()
    try:
        existing = conn.execute("SELECT id FROM decision_trees WHERE name = ?", (name,)).fetchone()
        if existing:
            conn.execute("UPDATE decision_trees SET category=?, description=?, start_node=?, nodes_json=?, updated_at=? WHERE id=?", (category, description, start_node, payload, now, existing["id"]))
        else:
            conn.execute("INSERT INTO decision_trees (name, category, description, start_node, nodes_json, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)", (name, category, description, start_node, payload, now, now))
        conn.commit(); return True, "Decision tree saved successfully."
    except sqlite3.IntegrityError:
        return False, "A decision tree with that name already exists."
    finally:
        conn.close()


def delete_decision_tree(tree_id):
    conn = db(); conn.execute("DELETE FROM decision_trees WHERE id = ?", (tree_id,)); conn.commit(); conn.close()


def parse_tree_nodes(raw_json):
    try:
        value = json.loads(raw_json)
        if not isinstance(value, dict): return None, "The tree definition must be a JSON object keyed by node ID."
        for node_id, node in value.items():
            if not isinstance(node, dict): return None, f"Node '{node_id}' must be an object."
            typ = str(node.get("type", "question")).lower()
            if typ not in {"question", "action"}: return None, f"Node '{node_id}' must have type 'question' or 'action'."
            if typ == "question":
                if not str(node.get("question", "")).strip(): return None, f"Question node '{node_id}' needs a question."
                answers = node.get("answers", [])
                if not isinstance(answers, list) or not answers: return None, f"Question node '{node_id}' needs at least one answer."
                for answer in answers:
                    if not isinstance(answer, dict) or not str(answer.get("label", "")).strip() or not str(answer.get("next", "")).strip():
                        return None, f"Every answer in '{node_id}' needs a label and next node."
                    if str(answer.get("next")) not in value:
                        return None, f"Answer '{answer.get('label')}' in '{node_id}' points to missing node '{answer.get('next')}'."
        return value, None
    except json.JSONDecodeError as e:
        return None, f"Invalid JSON: {e}"


def decision_tree_answer(label, next_node):
    st.session_state.decision_tree_history.append({"question": st.session_state.decision_tree_current_question, "answer": label})
    st.session_state.decision_tree_current_node = next_node
    st.rerun()


def render_decision_tree_page():
    st.markdown("<div class=\"page-heading\"><div><div class=\"page-title\">Decision Tree</div><div class=\"page-description\">Answer each question and the next action is determined by your response.</div></div></div>", unsafe_allow_html=True)
    trees = get_decision_trees()
    if not trees:
        st.info("No decision trees have been configured yet. An administrator can create one from the gear menu.")
        return
    options = {t["name"]: t["id"] for t in trees}
    selected = st.selectbox("Select a decision tree", list(options), key="decision_tree_selector")
    tree = get_decision_tree(options[selected])
    if not tree:
        st.error("The selected decision tree could not be loaded."); return
    try:
        nodes = json.loads(tree["nodes_json"])
    except Exception:
        st.error("This decision tree contains an invalid definition."); return
    if st.session_state.get("decision_tree_id") != tree["id"]:
        st.session_state.decision_tree_id = tree["id"]
        st.session_state.decision_tree_current_node = tree["start_node"]
        st.session_state.decision_tree_history = []
    current = st.session_state.get("decision_tree_current_node", tree["start_node"])
    node = nodes.get(current)
    if not node:
        st.error(f"Decision tree node '{current}' was not found."); return
    if st.session_state.get("decision_tree_history"):
        st.markdown("**Path:** " + " → ".join(x["answer"] for x in st.session_state.decision_tree_history))
    if str(node.get("type", "question")).lower() == "action":
        st.success("Next Action")
        st.markdown(f"### {node.get('action', 'No action specified.')}")
        if node.get("instructions"): st.markdown(node["instructions"])
        if node.get("source_pdf"): st.caption(f"Source PDF: {node['source_pdf']}" + (f" · Page {node['source_page']}" if node.get("source_page") else ""))
        if st.button("↻ Start Over", type="primary"):
            st.session_state.decision_tree_current_node = tree["start_node"]; st.session_state.decision_tree_history = []; st.rerun()
        return
    st.markdown(f"<div class='best-match-card'><h3>{node.get('question', 'Question')}</h3><div>Choose the answer that matches the current situation.</div></div>", unsafe_allow_html=True)
    st.session_state.decision_tree_current_question = node.get("question", "Question")
    for i, answer in enumerate(node.get("answers", [])):
        label = str(answer.get("label", "Answer")); nxt = str(answer.get("next", "")).strip()
        if st.button(label, key=f"decision_answer_{tree['id']}_{current}_{i}", use_container_width=True): decision_tree_answer(label, nxt)
    if st.button("↻ Restart", key=f"decision_restart_{tree['id']}"):
        st.session_state.decision_tree_current_node = tree["start_node"]; st.session_state.decision_tree_history = []; st.rerun()


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

if "decision_tree_id" not in st.session_state: st.session_state.decision_tree_id = None
if "decision_tree_current_node" not in st.session_state: st.session_state.decision_tree_current_node = None
if "decision_tree_current_question" not in st.session_state: st.session_state.decision_tree_current_question = ""
if "decision_tree_history" not in st.session_state: st.session_state.decision_tree_history = []


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


# HEADER
# ============================================================

user_name = "Authorized User"

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
            <div class="app-subtitle">Find exact information from your organization's documents</div>
        </div>
        <div class="top-user">{user_name}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Gear/admin menu.
# Keep the gear in a dedicated keyed wrapper so the Filters popover is never repositioned.
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

            if st.button("Manage Decision Trees", use_container_width=True):
                st.session_state.page = "Manage Decision Trees"
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
# ADMIN: MANAGE DECISION TREES
# ============================================================

elif st.session_state.page == "Manage Decision Trees":
    if not st.session_state.admin_authenticated:
        st.session_state.page = "Admin Login"; st.rerun()
    st.markdown("<div class=\"page-heading\"><div><div class=\"page-title\">Manage Decision Trees</div><div class=\"page-description\">Build guided workflows where each answer determines the next question or action.</div></div><div class=\"admin-badge\">ADMIN ONLY</div></div>", unsafe_allow_html=True)
    trees = get_decision_trees(); names = [t["name"] for t in trees]
    choice = st.selectbox("Tree", ["+ Create New"] + names, key="dt_edit_choice")
    selected = None if choice == "+ Create New" else next((t for t in trees if t["name"] == choice), None)
    if selected:
        name0, cat0, desc0, start0, json0 = selected["name"], selected["category"], selected["description"], selected["start_node"], selected["nodes_json"]
    else:
        name0, cat0, desc0, start0 = "", "General", "", "start"
        json0 = json.dumps({"start":{"type":"question","question":"Is the issue related to account access?","answers":[{"label":"Yes","next":"account_access"},{"label":"No","next":"not_account_access"}]},"account_access":{"type":"action","action":"Verify the user's account and follow the Account Access procedure."},"not_account_access":{"type":"action","action":"Continue with the applicable troubleshooting workflow."}}, indent=2)
    cats = ["General", "Policies", "Procedures", "Technical Support", "Licensing", "Training", "Product", "Account Management", "Other"]
    with st.form("decision_tree_form"):
        c1, c2 = st.columns([2, 1])
        with c1:
            name = st.text_input("Tree Name", value=name0, placeholder="Example: Licensing Troubleshooting")
            desc = st.text_input("Description", value=desc0, placeholder="When should this tree be used?")
        with c2:
            cat = st.selectbox("Category", cats, index=(cats.index(cat0) if cat0 in cats else 0))
            start_node = st.text_input("Start Node ID", value=start0)
        raw = st.text_area("Decision Tree Definition (JSON)", value=json0, height=480, help="Question nodes contain answers with next node IDs. Action nodes contain the next action.")
        submitted = st.form_submit_button("Save Decision Tree", type="primary", use_container_width=True)
    if submitted:
        nodes, error = parse_tree_nodes(raw)
        if error: st.error(error)
        else:
            ok, msg = save_decision_tree(name, cat, desc, start_node, nodes)
            if ok: st.success(msg); st.session_state.decision_tree_id = None; st.rerun()
            else: st.error(msg)
    if selected and st.button("Delete This Decision Tree"):
        delete_decision_tree(selected["id"]); st.session_state.decision_tree_id = None; st.rerun()
    with st.expander("Decision Tree Format Example"):
        st.code(json0, language="json")
        st.caption("Question nodes display answer buttons. Each answer points to another node. Action nodes end the path and display the next action. No AI model is used.")
    if st.button("← Back to Search"):
        st.session_state.page = "Search"; st.rerun()


# SEARCH KNOWLEDGE BASE
# ============================================================

elif st.session_state.page == "Decision Tree":
    render_decision_tree_page()

else:
    st.session_state.page = "Search"

    search_col, button_col, filter_col = st.columns([6.4, 1.0, 1.0], gap="small")
    with search_col:
        query = st.text_input(
            "Search",
            value=st.session_state.search_query,
            placeholder="What should I check or find in the knowledge base?",
            label_visibility="collapsed",
            key="main_search_box",
        )
    with button_col:
        search_clicked = st.button("Search", type="primary", use_container_width=True)
    with filter_col:
        with st.popover("☷ Filters", use_container_width=True):
            category = st.selectbox("Category", ["All Categories"] + get_categories(), key="search_category")
            top_k = st.selectbox("Results", [5, 10, 20], index=1, key="search_top_k")

    if search_clicked:
        st.session_state.search_query = query
        st.session_state.force_result_id = None
        st.session_state.search_signature = None

    active_query = st.session_state.search_query.strip()
    category = st.session_state.get("search_category", "All Categories")
    top_k = st.session_state.get("search_top_k", 10)

    if active_query:
        search_signature = (active_query, category, top_k)
        if st.session_state.get("search_signature") != search_signature:
            st.session_state.search_results = search_documents(active_query, category=category, top_k=top_k)
            st.session_state.search_signature = search_signature
            st.session_state.viewer_page = None

        results = st.session_state.get("search_results", [])

        if st.session_state.force_result_id is not None:
            forced_id = st.session_state.force_result_id
            forced = [r for r in results if r["id"] == forced_id]
            others = [r for r in results if r["id"] != forced_id]
            if forced:
                results = forced + others
                st.session_state.viewer_page = forced[0]["page_number"]
            st.session_state.force_result_id = None

        if not results:
            st.warning("No matching PDF was found. Try different keywords or upload another document.")
        else:
            best = results[0]
            best_doc = get_document_by_id(best["document_id"])
            total_pages = int(best_doc["page_count"] or 0) if best_doc else 0

            if st.session_state.get("viewer_page") is None:
                st.session_state.viewer_page = int(best["page_number"])

            viewer_page = max(1, min(int(st.session_state.viewer_page), max(total_pages, 1)))

            st.markdown(
                f'<div class="search-count">{len(results)} search result(s) · Showing the highest match first</div>',
                unsafe_allow_html=True,
            )

            related_col, source_col = st.columns([0.82, 1.8], gap="large")

            with related_col:
                st.markdown('<div class="panel-title">Search Results</div>', unsafe_allow_html=True)

                for i, result in enumerate(results):
                    is_best = i == 0
                    key_prefix = "best" if is_best else str(i)
                    with st.container(key=f"search_result_{key_prefix}", border=True):
                        score_pct = min(99, max(1, round(result["score"] * 100)))
                        label = "Highest Match" if is_best else "Search Result"
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
                            st.session_state.force_result_id = result["id"]
                            st.session_state.viewer_page = result["page_number"]
                            st.rerun()

                matches = find_document_matches(best["stored_path"], active_query, limit=8)
                st.markdown('<div class="match-panel"><div class="match-panel-title">Matches in this PDF</div></div>', unsafe_allow_html=True)
                if matches:
                    for match in matches:
                        if st.button(f"Page {match['page']}", key=f"jump_match_{best['id']}_{match['page']}", use_container_width=True):
                            st.session_state.viewer_page = match["page"]
                            st.rerun()
                        st.markdown(
                            f"<div class='match-item'><div class='match-item-page'>Page {match['page']} · {match['score']} term match(es)</div><div class='match-item-text'>{match['snippet']}</div></div>",
                            unsafe_allow_html=True,
                        )
                else:
                    st.caption("No exact text occurrence was detected on the other pages.")


            with source_col:
                st.markdown(
                    f"""
                    <div class='reader-toolbar'>
                        <div class='reader-title'>▣ {best['filename']}</div>
                        <div class='reader-meta'>Page {viewer_page} of {total_pages} · Search: “{active_query}”</div>
                        <div class='reader-match'>Highest match · source text shown exactly as it appears in the PDF</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                nav1, nav2, nav3, nav4 = st.columns([1, 1.1, 1, 1.1], gap="small")
                with nav1:
                    if st.button("‹ Previous", disabled=(viewer_page <= 1), use_container_width=True, key="pdf_prev"):
                        st.session_state.viewer_page = max(1, viewer_page - 1)
                        st.rerun()
                with nav2:
                    st.markdown(f"<div style='text-align:center;padding-top:8px;font-size:11px;color:#687b87;'>Page <b>{viewer_page}</b> / {total_pages}</div>", unsafe_allow_html=True)
                with nav3:
                    if st.button("Next ›", disabled=(viewer_page >= total_pages), use_container_width=True, key="pdf_next"):
                        st.session_state.viewer_page = min(total_pages, viewer_page + 1)
                        st.rerun()
                with nav4:
                    zoom = st.selectbox("Zoom", [125, 150, 175, 200], index=1, format_func=lambda x: f"{x}%", label_visibility="collapsed", key="pdf_zoom")

                image = render_pdf_page_highlighted(best["stored_path"], viewer_page, active_query, scale=zoom / 100 * 1.25)
                if image:
                    # Do not force the image to the column width: that makes every zoom
                    # level look identical.  A fixed display width lets the selected zoom
                    # visibly enlarge the rendered PDF page.
                    display_width = {125: 780, 150: 930, 175: 1080, 200: 1230}.get(int(zoom), 930)
                    st.image(image, width=display_width)
                else:
                    st.error("Unable to render this PDF page.")

                current_text = ""
                try:
                    with fitz.open(best["stored_path"]) as pdf:
                        current_text = clean_text(pdf[viewer_page - 1].get_text("text"))
                except Exception:
                    current_text = ""

                if current_text:
                    st.markdown(
                        f"""
                        <div class='match-panel'>
                            <div class='match-panel-title'>Match context · Page {viewer_page}</div>
                            <div class='result-snippet' style='font-size:12px;line-height:1.55;color:#294a5c;'>{make_snippet(current_text, active_query, radius=520)}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

    else:
        docs = get_documents()
        total_pages = sum(int(d["page_count"] or 0) for d in docs)
        st.markdown(
            f"""
            <div class="welcome-card">
                <div class="welcome-icon">⌕</div>
                <div class="welcome-title">Search your knowledge base</div>
                <div class="welcome-text">Search for a process, checklist, policy, troubleshooting step, or any information contained in your uploaded PDFs.</div>
                <div class="welcome-stats"><span><b>{len(docs)}</b> documents</span><span><b>{total_pages:,}</b> indexed pages</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ============================================================
# BOTTOM NAVIGATION
# ============================================================

st.markdown('<div class="bottom-nav-spacer"></div>', unsafe_allow_html=True)
nav_left, nav_center, nav_right = st.columns([1, 2, 1])
with nav_center:
    nav1, nav2 = st.columns(2, gap="small")
    with nav1:
        if st.button("⌕  Search Knowledge Base", use_container_width=True):
            st.session_state.page = "Search"; st.rerun()
    with nav2:
        if st.button("⑂  Decision Tree", use_container_width=True):
            st.session_state.page = "Decision Tree"; st.rerun()

st.markdown('<div class="bottom-nav-label">Knowledge Base · PDF Search · Guided Decision Trees</div>', unsafe_allow_html=True)
