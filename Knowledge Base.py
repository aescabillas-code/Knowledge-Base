"""
HPE Knowledge Base — Enhanced single-file Streamlit application

Core behavior
-------------
1. The user enters ACCESS_CODE once.
2. A signed browser token is issued and stored in the URL query parameter.
3. Refreshing the page does not ask for the access code again while the token
   remains valid.
4. The user can revoke the browser session from the settings menu.
5. Admin mode is entered from the same interface using ADMIN_PIN.
6. Admins can upload, replace, categorize, version, and delete PDFs.
7. PDFs are split into searchable page chunks and indexed with TF-IDF.
8. Search results show a direct extracted answer, supporting citations, and
   the original PDF page with matching text highlighted.
9. The PDF reader has its own page navigation and zoom controls; it does not
   depend on the browser's native PDF toolbar.

Recommended packages
--------------------
streamlit
PyMuPDF
pandas
scikit-learn
itsdangerous

Run
---
streamlit run "Knowledge Base.py"

Recommended Streamlit Secrets
-----------------------------
ACCESS_CODE = "..."
ADMIN_PIN = "..."
TOKEN_SECRET = "..."

Optional:
TOKEN_MAX_AGE_SECONDS = 28800
MAX_UPLOAD_MB = 50
"""

from __future__ import annotations

import hashlib
import html
import os
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

import fitz
import pandas as pd
import streamlit as st

try:
    from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
except Exception:
    BadSignature = Exception
    SignatureExpired = Exception
    URLSafeTimedSerializer = None

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ============================================================
# CONFIGURATION
# ============================================================

APP_NAME = "Knowledge Base"

# Fixed credentials requested by the user.
# For production, move these three values into Streamlit Secrets.
ADMIN_PIN = "HPE@123456"
ACCESS_CODE = "K8vQ2mR7xP4nZ9tL6wY3"
TOKEN_SECRET = "f7Xq9L2vN8mR4tY6pK3zW5cJ1sH8dQ0aV6eB2nG9xT4uP7"

TOKEN_MAX_AGE_SECONDS = int(
    st.secrets.get("TOKEN_MAX_AGE_SECONDS", 28800)
    if hasattr(st, "secrets")
    else 28800
)
MAX_UPLOAD_MB = int(
    st.secrets.get("MAX_UPLOAD_MB", 50)
    if hasattr(st, "secrets")
    else 50
)

# Secrets override hard-coded development values when supplied.
try:
    ADMIN_PIN = str(st.secrets.get("ADMIN_PIN", ADMIN_PIN)).strip()
    ACCESS_CODE = str(st.secrets.get("ACCESS_CODE", ACCESS_CODE)).strip()
    TOKEN_SECRET = str(st.secrets.get("TOKEN_SECRET", TOKEN_SECRET)).strip()
except Exception:
    pass

DATA_DIR = Path(os.getenv("KB_DATA_DIR", "knowledge_base_data"))
PDF_DIR = DATA_DIR / "pdfs"
DB_PATH = DATA_DIR / "knowledge_base.db"

DATA_DIR.mkdir(parents=True, exist_ok=True)
PDF_DIR.mkdir(parents=True, exist_ok=True)

CATEGORIES = [
    "General",
    "Account Management",
    "Licensing",
    "Portal & Access",
    "Technical Support",
    "HPE GreenLake",
    "Product Guides",
    "Policies & Procedures",
    "Troubleshooting",
]

CATEGORY_ICONS = {
    "General": "📚",
    "Account Management": "👤",
    "Licensing": "🔑",
    "Portal & Access": "🖥️",
    "Technical Support": "🛠️",
    "HPE GreenLake": "☁️",
    "Product Guides": "📖",
    "Policies & Procedures": "🛡️",
    "Troubleshooting": "⚙️",
}


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="HPE Knowledge Base",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
<style>
:root {
    --navy:#0b2538;
    --navy2:#123b50;
    --teal:#00a982;
    --teal-dark:#007f72;
    --soft:#e8f7f4;
    --bg:#f5f8fa;
    --white:#ffffff;
    --border:#d9e3e8;
    --text:#102d42;
    --muted:#687b87;
    --danger:#c83b3b;
    --gold:#c88900;
}

html, body, [class*="css"] {
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
}

.stApp {
    background:linear-gradient(180deg,#f8fbfc 0%,#f2f6f8 100%);
    color:var(--text);
}

[data-testid="stHeader"] {
    height:0 !important;
    min-height:0 !important;
    background:transparent !important;
}

[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"],
[data-testid="stMainMenu"],
[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"] {
    display:none !important;
    visibility:hidden !important;
}

footer { visibility:hidden !important; }

button[kind="primary"] {
    background:var(--teal) !important;
    border-color:var(--teal) !important;
    color:white !important;
}

button[kind="primary"]:hover {
    background:var(--teal-dark) !important;
    border-color:var(--teal-dark) !important;
}

.kb-topbar {
    border-bottom:1px solid #dfe7eb;
    background:linear-gradient(105deg,#ffffff 0%,#f7fbfc 70%,#e8f7f7 100%);
    border-radius:10px;
    padding:10px 18px;
    margin-bottom:16px;
    box-shadow:0 1px 4px rgba(0,0,0,.03);
}

.kb-brand {
    display:flex;
    align-items:center;
    gap:14px;
}

.kb-brand-mark {
    width:42px;
    height:6px;
    background:var(--teal);
    border-radius:2px;
    margin-bottom:5px;
}

.kb-brand-name {
    font-size:13px;
    font-weight:800;
    line-height:1.1;
    color:#111;
}

.kb-title {
    font-size:24px;
    font-weight:750;
    line-height:1.1;
    color:var(--text);
}

.kb-subtitle {
    font-size:12px;
    color:#304a5c;
    margin-top:3px;
}

.user-pill {
    display:inline-flex;
    align-items:center;
    justify-content:center;
    gap:5px;
    background:#e8f7f7;
    color:#315468;
    font-size:12px;
    padding:6px 12px;
    border-radius:20px;
    font-weight:700;
    border:1px solid #cce5df;
    white-space:nowrap;
}

.admin-pill {
    background:#fff4d7;
    color:#805800;
    border-color:#efd28b;
}

.hero {
    border-radius:14px;
    padding:28px 32px;
    margin:8px 0 18px;
    color:white;
    background:
        radial-gradient(circle at 88% 20%,rgba(0,169,130,.35),transparent 30%),
        linear-gradient(115deg,#073b40,#006b70 58%,#008f83);
    box-shadow:0 10px 30px rgba(3,63,67,.12);
}

.hero h1 {
    margin:0 0 7px;
    color:white;
    font-size:31px;
}

.hero p {
    margin:0 0 18px;
    color:#d7f1f1;
    font-size:14px;
}

.hero-chip {
    display:inline-block;
    padding:5px 10px;
    margin:4px 4px 0 0;
    background:rgba(255,255,255,.12);
    border:1px solid rgba(255,255,255,.15);
    border-radius:18px;
    color:white;
    font-size:11px;
}

.section-title {
    font-size:19px;
    font-weight:800;
    color:var(--text);
    margin:16px 0 10px;
}

.metric-card {
    background:#fff;
    border:1px solid #e1eaee;
    border-radius:13px;
    padding:16px;
    min-height:105px;
    box-shadow:0 3px 12px rgba(16,42,67,.045);
}

.metric-icon {
    width:40px;
    height:40px;
    display:inline-flex;
    align-items:center;
    justify-content:center;
    border-radius:11px;
    background:#e7f8f1;
    font-size:21px;
}

.metric-value {
    font-size:25px;
    font-weight:850;
    color:var(--text);
    margin-top:4px;
}

.metric-label {
    color:#6b8290;
    font-size:12px;
}

.category-card {
    background:white;
    border:1px solid #e1eaee;
    border-radius:12px;
    padding:14px;
    min-height:78px;
    box-shadow:0 3px 10px rgba(16,42,67,.035);
}

.category-icon {
    float:left;
    font-size:22px;
    margin-right:10px;
}

.category-title {
    color:var(--text);
    font-weight:750;
    font-size:13px;
}

.category-count {
    color:#8296a3;
    font-size:11px;
}

.doc-card {
    background:white;
    border:1px solid #e1eaee;
    border-radius:12px;
    padding:13px;
    margin-bottom:8px;
    box-shadow:0 2px 9px rgba(16,42,67,.03);
}

.pdf-badge {
    display:inline-flex;
    width:38px;
    height:38px;
    align-items:center;
    justify-content:center;
    border-radius:9px;
    background:#fff0ef;
    color:#d64545;
    font-size:11px;
    font-weight:850;
    margin-right:9px;
}

.exact-answer {
    background:linear-gradient(110deg,#f0fdf4 0%,#fff 85%);
    border:1px solid #86efac;
    border-left:5px solid var(--teal);
    border-radius:9px;
    padding:14px 17px;
    margin-bottom:14px;
    box-shadow:0 2px 8px rgba(0,169,130,.05);
}

.answer-label {
    color:#047857;
    font-size:11px;
    font-weight:850;
    text-transform:uppercase;
    letter-spacing:.5px;
}

.answer-pill {
    background:#dcfce7;
    color:#15803d;
    padding:3px 9px;
    border-radius:12px;
    font-size:10px;
    font-weight:800;
}

.answer-text {
    color:#0f172a;
    font-size:14px;
    line-height:1.5;
    font-weight:500;
    margin:8px 0;
}

.answer-meta {
    color:#64748b;
    font-size:10px;
}

.reader-toolbar {
    background:#fff;
    border:1px solid var(--border);
    border-radius:8px 8px 0 0;
    padding:9px 13px;
}

.reader-title {
    color:var(--text);
    font-size:13px;
    font-weight:800;
    word-break:break-word;
}

.reader-meta {
    color:var(--muted);
    font-size:10px;
    margin-top:2px;
}

.result-snippet {
    color:#475569;
    font-size:11.5px;
    line-height:1.45;
    margin-top:5px;
}

.result-label {
    font-size:10px;
    font-weight:800;
    text-transform:uppercase;
    letter-spacing:.4px;
}

.welcome-card {
    margin:45px auto;
    max-width:700px;
    text-align:center;
    background:white;
    border:1px solid var(--border);
    border-radius:14px;
    padding:36px;
    box-shadow:0 4px 18px rgba(12,54,70,.04);
}

.welcome-icon {
    font-size:38px;
    color:var(--teal);
}

.welcome-title {
    font-size:22px;
    font-weight:800;
    color:var(--text);
    margin-top:7px;
}

.welcome-text {
    color:var(--muted);
    max-width:530px;
    margin:8px auto;
    line-height:1.5;
    font-size:13px;
}

.auth-shell {
    max-width:540px;
    margin:80px auto 18px;
    text-align:center;
}

.auth-mark {
    width:48px;
    height:6px;
    background:var(--teal);
    margin:0 auto 9px;
    border-radius:2px;
}

.auth-brand {
    font-size:13px;
    font-weight:800;
    color:#111;
}

.auth-title {
    margin-top:18px;
    font-size:29px;
    font-weight:800;
    color:var(--text);
}

.auth-subtitle {
    margin-top:6px;
    color:var(--muted);
    font-size:13px;
    margin-bottom:22px;
}

.admin-banner {
    background:linear-gradient(100deg,#062f34,#087d77);
    color:white;
    border-radius:11px;
    padding:13px 16px;
    margin-bottom:14px;
}

.admin-banner strong { color:white; }

div[data-testid="stPopover"] > button {
    border-radius:7px !important;
    border:1px solid #cce5df !important;
    background:#fff !important;
    height:34px !important;
    min-height:34px !important;
}

@media (max-width: 900px) {
    .kb-title { font-size:20px; }
    .kb-subtitle { display:none; }
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# DATABASE
# ============================================================

def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = db()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA foreign_keys=ON")

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            stored_path TEXT NOT NULL,
            file_hash TEXT NOT NULL UNIQUE,
            title TEXT NOT NULL,
            category TEXT NOT NULL DEFAULT 'General',
            description TEXT DEFAULT '',
            tags TEXT DEFAULT '',
            page_count INTEGER DEFAULT 0,
            file_size INTEGER DEFAULT 0,
            version TEXT DEFAULT '1.0',
            uploaded_by TEXT DEFAULT '',
            uploaded_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            status TEXT DEFAULT 'Indexed',
            view_count INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS chunks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER NOT NULL,
            page_number INTEGER NOT NULL,
            chunk_index INTEGER NOT NULL,
            text TEXT NOT NULL,
            FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_type TEXT NOT NULL,
            document_id INTEGER,
            created_at TEXT NOT NULL,
            details TEXT DEFAULT ''
        );

        CREATE INDEX IF NOT EXISTS idx_chunks_document_page
            ON chunks(document_id, page_number);

        CREATE INDEX IF NOT EXISTS idx_chunks_document
            ON chunks(document_id);

        CREATE INDEX IF NOT EXISTS idx_documents_category
            ON documents(category);
        """
    )
    conn.commit()
    conn.close()


init_db()


# ============================================================
# AUTHENTICATION
# ============================================================

def token_serializer():
    if not TOKEN_SECRET or URLSafeTimedSerializer is None:
        return None
    return URLSafeTimedSerializer(
        TOKEN_SECRET,
        salt="hpe-knowledge-base-access-v2",
    )


def create_browser_token() -> str:
    serializer = token_serializer()
    if serializer is None:
        return ""
    return serializer.dumps(
        {
            "authorized": True,
            "issued_for": "knowledge-base",
        }
    )


def validate_browser_token(token: str) -> bool:
    if not token:
        return False

    serializer = token_serializer()
    if serializer is None:
        return False

    try:
        payload = serializer.loads(
            str(token),
            max_age=TOKEN_MAX_AGE_SECONDS,
        )
        return bool(payload.get("authorized"))
    except (BadSignature, SignatureExpired, Exception):
        return False


def browser_is_authorized() -> bool:
    if st.session_state.get("access_authorized", False):
        return True

    token = st.query_params.get("kb_access", "")

    if validate_browser_token(token):
        st.session_state.access_authorized = True
        return True

    return False


def authorize_browser() -> bool:
    token = create_browser_token()
    if not token:
        return False

    st.query_params["kb_access"] = token
    st.session_state.access_authorized = True
    return True


def revoke_browser_access() -> None:
    st.session_state.access_authorized = False
    st.session_state.admin_authenticated = False

    try:
        st.query_params.clear()
    except Exception:
        pass


def admin_is_authorized() -> bool:
    return bool(st.session_state.get("admin_authenticated", False))


# ============================================================
# STATE
# ============================================================

DEFAULT_STATE = {
    "access_authorized": False,
    "admin_authenticated": False,
    "page": "Home",
    "search_query": "",
    "search_signature": None,
    "search_results": [],
    "selected_document": None,
    "viewer_page": 1,
    "highlight_target": "",
    "new_access_message": "",
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# UTILITIES
# ============================================================

def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def clean_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def file_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def safe_filename(name: str) -> str:
    return re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        Path(name).name,
    ).strip("_") or "document.pdf"


def format_bytes(value: int) -> str:
    n = float(value or 0)

    for unit in ["B", "KB", "MB", "GB"]:
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024

    return f"{n:.1f} TB"


def split_text(text: str, chunk_size: int = 850, overlap: int = 120) -> list[str]:
    words = text.split()

    if not words:
        return []

    chunks: list[str] = []
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


def extract_pdf(pdf_bytes: bytes) -> tuple[list[tuple[int, str]], int]:
    pages = []

    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        for page_number, page in enumerate(document, start=1):
            pages.append(
                (
                    page_number,
                    clean_text(page.get_text("text")),
                )
            )

        return pages, len(document)


def save_pdf(filename: str, pdf_bytes: bytes, digest: str) -> str:
    destination = PDF_DIR / f"{digest[:12]}_{safe_filename(filename)}"
    destination.write_bytes(pdf_bytes)
    return str(destination)


def record_event(
    event_type: str,
    document_id: Optional[int] = None,
    details: str = "",
) -> None:
    conn = db()
    conn.execute(
        """
        INSERT INTO events(event_type, document_id, created_at, details)
        VALUES (?, ?, ?, ?)
        """,
        (
            event_type,
            document_id,
            now_iso(),
            details,
        ),
    )
    conn.commit()
    conn.close()


# ============================================================
# DOCUMENT REPOSITORY
# ============================================================

def get_documents() -> list[dict[str, Any]]:
    conn = db()

    rows = conn.execute(
        """
        SELECT *
        FROM documents
        ORDER BY updated_at DESC
        """
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


def get_document(document_id: int) -> Optional[dict[str, Any]]:
    conn = db()

    row = conn.execute(
        "SELECT * FROM documents WHERE id = ?",
        (document_id,),
    ).fetchone()

    conn.close()

    return dict(row) if row else None


def get_categories() -> list[str]:
    conn = db()

    rows = conn.execute(
        """
        SELECT DISTINCT category
        FROM documents
        ORDER BY category
        """
    ).fetchall()

    conn.close()

    return [row["category"] for row in rows]


def get_all_chunks() -> list[dict[str, Any]]:
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
            d.title,
            d.category,
            d.stored_path,
            d.version
        FROM chunks c
        JOIN documents d ON d.id = c.document_id
        ORDER BY c.id
        """
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


def add_document(
    filename: str,
    pdf_bytes: bytes,
    category: str,
    title: str,
    description: str,
    tags: str,
    version: str,
    uploaded_by: str,
) -> tuple[bool, str]:
    digest = file_hash(pdf_bytes)

    conn = db()

    exists = conn.execute(
        "SELECT id FROM documents WHERE file_hash = ?",
        (digest,),
    ).fetchone()

    if exists:
        conn.close()
        return False, "This exact PDF already exists in the knowledge base."

    try:
        pages, page_count = extract_pdf(pdf_bytes)

        if page_count == 0:
            conn.close()
            return False, "The PDF contains no pages."

        stored_path = save_pdf(filename, pdf_bytes, digest)
        timestamp = now_iso()

        cursor = conn.execute(
            """
            INSERT INTO documents (
                filename,
                stored_path,
                file_hash,
                title,
                category,
                description,
                tags,
                page_count,
                file_size,
                version,
                uploaded_by,
                uploaded_at,
                updated_at,
                status
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                filename,
                stored_path,
                digest,
                title or Path(filename).stem,
                category,
                description,
                tags,
                page_count,
                len(pdf_bytes),
                version or "1.0",
                uploaded_by,
                timestamp,
                timestamp,
                "Indexed",
            ),
        )

        document_id = cursor.lastrowid

        for page_number, page_text in pages:
            for chunk_index, chunk in enumerate(
                split_text(page_text)
            ):
                conn.execute(
                    """
                    INSERT INTO chunks (
                        document_id,
                        page_number,
                        chunk_index,
                        text
                    )
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

    except Exception as exc:
        conn.rollback()
        try:
            Path(stored_path).unlink(missing_ok=True)
        except Exception:
            pass
        conn.close()
        return False, f"PDF indexing failed: {exc}"

    conn.close()

    record_event(
        "upload",
        int(document_id),
        filename,
    )

    return True, f"Indexed {filename} — {page_count} pages."


def replace_document(
    document_id: int,
    pdf_bytes: bytes,
    filename: str,
    category: str,
    title: str,
    description: str,
    tags: str,
    version: str,
    uploaded_by: str,
) -> tuple[bool, str]:
    digest = file_hash(pdf_bytes)

    conn = db()

    existing = conn.execute(
        "SELECT * FROM documents WHERE id = ?",
        (document_id,),
    ).fetchone()

    if not existing:
        conn.close()
        return False, "Document no longer exists."

    duplicate = conn.execute(
        """
        SELECT id
        FROM documents
        WHERE file_hash = ? AND id != ?
        """,
        (digest, document_id),
    ).fetchone()

    if duplicate:
        conn.close()
        return False, "That PDF already exists as another document."

    old_path = existing["stored_path"]

    try:
        pages, page_count = extract_pdf(pdf_bytes)
        stored_path = save_pdf(filename, pdf_bytes, digest)
        timestamp = now_iso()

        conn.execute(
            """
            UPDATE documents
            SET filename = ?,
                stored_path = ?,
                file_hash = ?,
                title = ?,
                category = ?,
                description = ?,
                tags = ?,
                page_count = ?,
                file_size = ?,
                version = ?,
                uploaded_by = ?,
                updated_at = ?,
                status = 'Indexed'
            WHERE id = ?
            """,
            (
                filename,
                stored_path,
                digest,
                title or Path(filename).stem,
                category,
                description,
                tags,
                page_count,
                len(pdf_bytes),
                version or "1.0",
                uploaded_by,
                timestamp,
                document_id,
            ),
        )

        conn.execute(
            "DELETE FROM chunks WHERE document_id = ?",
            (document_id,),
        )

        for page_number, page_text in pages:
            for chunk_index, chunk in enumerate(
                split_text(page_text)
            ):
                conn.execute(
                    """
                    INSERT INTO chunks (
                        document_id,
                        page_number,
                        chunk_index,
                        text
                    )
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

    except Exception as exc:
        conn.rollback()
        try:
            Path(stored_path).unlink(missing_ok=True)
        except Exception:
            pass
        conn.close()
        return False, f"Replacement failed: {exc}"

    conn.close()

    try:
        Path(old_path).unlink(missing_ok=True)
    except Exception:
        pass

    record_event(
        "replace",
        document_id,
        filename,
    )

    return True, f"Updated {filename} — {page_count} pages re-indexed."


def delete_document(document_id: int) -> None:
    conn = db()

    row = conn.execute(
        "SELECT stored_path, filename FROM documents WHERE id = ?",
        (document_id,),
    ).fetchone()

    if not row:
        conn.close()
        return

    path = row["stored_path"]
    filename = row["filename"]

    conn.execute(
        "DELETE FROM documents WHERE id = ?",
        (document_id,),
    )

    conn.commit()
    conn.close()

    try:
        Path(path).unlink(missing_ok=True)
    except Exception:
        pass

    record_event(
        "delete",
        document_id,
        filename,
    )


def increment_view(document_id: int) -> None:
    conn = db()

    conn.execute(
        """
        UPDATE documents
        SET view_count = view_count + 1
        WHERE id = ?
        """,
        (document_id,),
    )

    conn.commit()
    conn.close()

    record_event(
        "view",
        document_id,
    )


# ============================================================
# SEARCH ENGINE
# ============================================================

@st.cache_data(ttl=120, show_spinner=False)
def build_search_index(index_version: int):
    rows = get_all_chunks()

    if not rows:
        return None, None, []

    texts = [row["text"] for row in rows]

    vectorizer = TfidfVectorizer(
        lowercase=True,
        stop_words="english",
        ngram_range=(1, 2),
        sublinear_tf=True,
        min_df=1,
        dtype="float32",
    )

    matrix = vectorizer.fit_transform(texts)

    return vectorizer, matrix, rows


def current_index_version() -> int:
    conn = db()

    row = conn.execute(
        """
        SELECT
            COALESCE(MAX(id), 0) AS max_id,
            COALESCE(SUM(page_count), 0) AS pages
        FROM documents
        """
    ).fetchone()

    conn.close()

    return int(row["max_id"] * 100000 + row["pages"])


def make_snippet(text: str, query: str, radius: int = 230) -> str:
    clean = " ".join(text.split())

    terms = [
        token.lower()
        for token in re.findall(r"\w+", query)
        if len(token) > 2
    ]

    if not terms:
        return clean[:radius] + (
            "..." if len(clean) > radius else ""
        )

    lower = clean.lower()

    positions = [
        lower.find(term)
        for term in terms
        if lower.find(term) >= 0
    ]

    if not positions:
        return clean[:radius] + (
            "..." if len(clean) > radius else ""
        )

    start = max(0, min(positions) - 70)
    end = min(len(clean), start + radius)

    return (
        ("..." if start > 0 else "")
        + clean[start:end]
        + ("..." if end < len(clean) else "")
    )


def extract_direct_sentence(text: str, query: str) -> str:
    sentences = re.split(
        r"(?<=[.!?])\s+",
        clean_text(text),
    )

    terms = [
        token.lower()
        for token in re.findall(r"\w+", query)
        if len(token) > 2
    ]

    if not sentences:
        return text[:300]

    best = sentences[0]
    best_score = -1

    for sentence in sentences:
        lower = sentence.lower()
        score = sum(lower.count(term) for term in terms)

        if score > best_score and len(sentence.split()) >= 4:
            best = sentence.strip()
            best_score = score

    return best


@st.cache_data(ttl=120, show_spinner=False)
def search_documents(
    query: str,
    category: str,
    top_k: int,
    index_version: int,
) -> list[dict[str, Any]]:
    query = query.strip()

    if not query:
        return []

    vectorizer, matrix, rows = build_search_index(index_version)

    if vectorizer is None:
        return []

    query_vector = vectorizer.transform([query])
    scores = cosine_similarity(
        query_vector,
        matrix,
    ).flatten()

    results = []

    for index, score in enumerate(scores):
        if score <= 0.015:
            continue

        row = rows[index]

        if (
            category != "All Categories"
            and row["category"] != category
        ):
            continue

        results.append(
            {
                **row,
                "raw_score": float(score),
                "snippet": make_snippet(
                    row["text"],
                    query,
                ),
                "direct_sentence": extract_direct_sentence(
                    row["text"],
                    query,
                ),
            }
        )

    results.sort(
        key=lambda result: result["raw_score"],
        reverse=True,
    )

    return results[:top_k]


# ============================================================
# PDF READER
# ============================================================

@st.cache_data(ttl=300, show_spinner=False)
def render_pdf_page(
    path: str,
    page_number: int,
    target_text: str,
    zoom: int,
) -> Optional[bytes]:
    try:
        with fitz.open(path) as document:
            if not document:
                return None

            page_index = max(
                0,
                min(page_number - 1, len(document) - 1),
            )

            page = document[page_index]

            target = " ".join(
                target_text.split()[:14]
            )

            rectangles = []

            if target:
                rectangles = page.search_for(target)

            if not rectangles:
                words = target.split()

                # Progressive shorter search phrases.
                for length in [10, 7, 5, 3]:
                    if len(words) >= length:
                        phrase = " ".join(words[:length])
                        rectangles = page.search_for(phrase)

                        if rectangles:
                            break

            for rectangle in rectangles:
                annotation = page.add_highlight_annot(
                    rectangle
                )
                annotation.set_colors(
                    stroke=(0.0, 0.66, 0.51)
                )
                annotation.update()

            scale = max(0.75, zoom / 100.0) * 1.35

            pixmap = page.get_pixmap(
                matrix=fitz.Matrix(scale, scale),
                alpha=False,
            )

            return pixmap.tobytes("png")

    except Exception:
        return None


def clear_search_cache() -> None:
    try:
        build_search_index.clear()
        search_documents.clear()
        render_pdf_page.clear()
    except Exception:
        pass


# ============================================================
# TOPBAR
# ============================================================

def render_topbar() -> None:
    with st.container():
        st.markdown(
            """
            <div class="kb-topbar">
                <div class="kb-brand">
                    <div style="width:145px;">
                        <div class="kb-brand-mark"></div>
                        <div class="kb-brand-name">
                            HEWLETT PACKARD<br>ENTERPRISE
                        </div>
                    </div>

                    <div style="height:40px;width:1px;background:#b8c7cf;"></div>

                    <div>
                        <div class="kb-title">Knowledge Base</div>
                        <div class="kb-subtitle">
                            Find exact information from your organization's documents
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    left, right = st.columns([6.2, 1.8], vertical_alignment="center")

    with left:
        pass

    with right:
        label = (
            "🛡️ Admin Mode"
            if admin_is_authorized()
            else "👤 Authorized User"
        )

        pill_class = (
            "user-pill admin-pill"
            if admin_is_authorized()
            else "user-pill"
        )

        st.markdown(
            f'<div style="text-align:right;"><span class="{pill_class}">'
            f'{label}</span></div>',
            unsafe_allow_html=True,
        )

        with st.popover("⚙", use_container_width=True):
            st.markdown("**Workspace Options**")
            st.caption("Browser session is authenticated.")
            st.divider()

            if admin_is_authorized():
                st.success("Administrator mode is active.")

                if st.button(
                    "📚 Knowledge Base",
                    use_container_width=True,
                ):
                    st.session_state.page = "Home"
                    st.rerun()

                if st.button(
                    "🗂 Document Center",
                    use_container_width=True,
                ):
                    st.session_state.page = "Manage Documents"
                    st.rerun()

                if st.button(
                    "🔒 Lock Admin Mode",
                    use_container_width=True,
                ):
                    st.session_state.admin_authenticated = False
                    st.session_state.page = "Home"
                    st.rerun()

            else:
                if st.button(
                    "🛡️ Admin Control Panel",
                    use_container_width=True,
                ):
                    st.session_state.page = "Admin Login"
                    st.rerun()

            st.divider()

            if st.button(
                "↪ Revoke Browser Access",
                use_container_width=True,
            ):
                revoke_browser_access()
                st.rerun()


# ============================================================
# ACCESS GATE
# ============================================================

def render_access_gate() -> None:
    st.markdown(
        """
        <div class="auth-shell">
            <div class="auth-mark"></div>
            <div class="auth-brand">HEWLETT PACKARD ENTERPRISE</div>
            <div class="auth-title">Knowledge Base</div>
            <div class="auth-subtitle">
                Enter the access code once to authenticate this browser.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _, center, _ = st.columns([1, 1.35, 1])

    with center:
        with st.form("access_code_form"):
            code = st.text_input(
                "Access Code",
                type="password",
                placeholder="Enter access code",
            )

            submitted = st.form_submit_button(
                "Authenticate Workspace",
                type="primary",
                use_container_width=True,
            )

        if submitted:
            if code.strip() == ACCESS_CODE:
                if authorize_browser():
                    st.rerun()
                else:
                    st.error(
                        "Unable to issue a secure browser token."
                    )
            else:
                st.error("Invalid access code.")


if not browser_is_authorized():
    render_access_gate()
    st.stop()


# ============================================================
# HOME / DASHBOARD
# ============================================================

def metric_cards(docs: list[dict[str, Any]]) -> None:
    categories = len(
        set(document["category"] for document in docs)
    ) if docs else 0

    total_pages = sum(
        document["page_count"]
        for document in docs
    )

    total_views = sum(
        document["view_count"]
        for document in docs
    )

    cards = [
        ("📚", len(docs), "Total Documents"),
        ("📑", total_pages, "Searchable Pages"),
        ("📁", categories, "Categories"),
        ("👁️", total_views, "Document Views"),
    ]

    columns = st.columns(4)

    for column, (icon, value, label) in zip(
        columns,
        cards,
    ):
        with column:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-icon">{icon}</div>
                    <div class="metric-value">{value:,}</div>
                    <div class="metric-label">{label}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def category_cards(docs: list[dict[str, Any]]) -> None:
    st.markdown(
        '<div class="section-title">▦ Browse by Category</div>',
        unsafe_allow_html=True,
    )

    counts: dict[str, int] = {}

    for document in docs:
        category = document["category"]
        counts[category] = counts.get(category, 0) + 1

    categories = [
        category
        for category in CATEGORIES
        if counts.get(category, 0)
    ]

    if not categories:
        st.info(
            "No documents are indexed yet."
        )
        return

    columns = st.columns(4)

    for index, category in enumerate(categories):
        with columns[index % 4]:
            st.markdown(
                f"""
                <div class="category-card">
                    <span class="category-icon">
                        {CATEGORY_ICONS.get(category, "📄")}
                    </span>
                    <div class="category-title">
                        {html.escape(category)}
                    </div>
                    <div class="category-count">
                        {counts[category]} documents
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if st.button(
                "Open",
                key=f"cat_open_{index}_{category}",
                use_container_width=True,
            ):
                st.session_state.page = "Search"
                st.session_state.category_filter = category
                st.rerun()


def render_document_card(
    document: dict[str, Any],
    key_prefix: str,
) -> None:
    st.markdown(
        f"""
        <div class="doc-card">
            <span class="pdf-badge">PDF</span>
            <span style="font-weight:800;color:#102d42;">
                {html.escape(document["title"])}
            </span>
            <div style="font-size:11px;color:#8296a3;margin-top:3px;">
                {html.escape(document["category"])}
                &nbsp;•&nbsp; v{html.escape(document["version"])}
                &nbsp;•&nbsp; {document["page_count"]} pages
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    open_col, download_col, spacer = st.columns(
        [1, 1, 5]
    )

    with open_col:
        if st.button(
            "Open",
            key=f"{key_prefix}_open_{document['id']}",
            use_container_width=True,
        ):
            st.session_state.selected_document = document["id"]
            st.session_state.viewer_page = 1
            st.session_state.highlight_target = ""
            st.session_state.page = "Document"
            increment_view(document["id"])
            st.rerun()

    with download_col:
        path = Path(document["stored_path"])

        if path.exists():
            st.download_button(
                "Download",
                data=path.read_bytes(),
                file_name=document["filename"],
                mime="application/pdf",
                key=f"{key_prefix}_download_{document['id']}",
                use_container_width=True,
            )


def home_page() -> None:
    docs = get_documents()

    st.markdown(
        """
        <div class="hero">
            <h1>Find the answers you need</h1>
            <p>
                Search approved documentation, explore topics,
                or browse the knowledge repository.
            </p>
            <span class="hero-chip">Licensing</span>
            <span class="hero-chip">Portal Access</span>
            <span class="hero-chip">Account Setup</span>
            <span class="hero-chip">Troubleshooting</span>
            <span class="hero-chip">HPE GreenLake</span>
            <span class="hero-chip">Software Support</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    query = st.text_input(
        "Search",
        placeholder="Search for solutions, guides, error codes, procedures, or keywords...",
        label_visibility="collapsed",
        key="home_search",
    )

    if st.button(
        "Search Knowledge Base",
        type="primary",
        use_container_width=True,
    ):
        if query.strip():
            st.session_state.search_query = query.strip()
            st.session_state.page = "Search"
            st.session_state.search_signature = None
            st.rerun()

    metric_cards(docs)

    category_cards(docs)

    st.markdown(
        '<div class="section-title">📄 Recently Updated</div>',
        unsafe_allow_html=True,
    )

    if not docs:
        st.markdown(
            """
            <div class="welcome-card">
                <div class="welcome-icon">📚</div>
                <div class="welcome-title">
                    Your Knowledge Base is ready
                </div>
                <div class="welcome-text">
                    There are no PDFs indexed yet.
                    Administrators can open the Document Center
                    to upload approved documents.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    for document in docs[:5]:
        render_document_card(
            document,
            "home",
        )


# ============================================================
# SEARCH PAGE
# ============================================================

def search_page() -> None:
    query_col, button_col, filter_col = st.columns(
        [6.1, 1.1, 1.2],
        gap="small",
    )

    with query_col:
        query = st.text_input(
            "Search Knowledge Base",
            value=st.session_state.get(
                "search_query",
                "",
            ),
            placeholder=(
                "Type a procedure, error code, SKU, "
                "policy, or question..."
            ),
            label_visibility="collapsed",
        )

    with button_col:
        search_clicked = st.button(
            "Search",
            type="primary",
            use_container_width=True,
        )

    with filter_col:
        categories = [
            "All Categories"
        ] + get_categories()

        current_category = st.session_state.get(
            "category_filter",
            "All Categories",
        )

        if current_category not in categories:
            current_category = "All Categories"

        category = st.selectbox(
            "Category",
            categories,
            index=categories.index(current_category),
            label_visibility="collapsed",
        )

    if search_clicked:
        st.session_state.search_query = query.strip()
        st.session_state.category_filter = category
        st.session_state.search_signature = None
        st.session_state.selected_document = None
        st.session_state.viewer_page = 1
        st.session_state.highlight_target = ""
        st.rerun()

    active_query = st.session_state.get(
        "search_query",
        "",
    ).strip()

    active_category = st.session_state.get(
        "category_filter",
        category,
    )

    if not active_query:
        docs = get_documents()
        total_pages = sum(
            document["page_count"]
            for document in docs
        )

        st.markdown(
            f"""
            <div class="welcome-card">
                <div class="welcome-icon">⌕</div>
                <div class="welcome-title">
                    Search Organization Knowledge Base
                </div>
                <div class="welcome-text">
                    Search exact procedures, product information,
                    troubleshooting instructions, licensing details,
                    and policy content. Results are linked directly
                    to the source PDF page.
                </div>
                <div style="display:flex;justify-content:center;gap:30px;
                            color:#57707e;margin-top:16px;font-size:12px;">
                    <span><b>{len(docs)}</b> Documents Indexed</span>
                    <span><b>{total_pages:,}</b> Searchable Pages</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    signature = (
        active_query,
        active_category,
        12,
        current_index_version(),
    )

    if st.session_state.get("search_signature") != signature:
        st.session_state.search_results = search_documents(
            active_query,
            active_category,
            12,
            signature[-1],
        )
        st.session_state.search_signature = signature
        st.session_state.viewer_page = None
        st.session_state.highlight_target = ""
        st.session_state.selected_document = None

    results = st.session_state.get(
        "search_results",
        [],
    )

    if not results:
        st.warning(
            "No matching references found. "
            "Try a product name, error code, or a more specific phrase."
        )
        return

    # If a citation was clicked, move it to the top.
    selected_id = st.session_state.get(
        "selected_document",
    )

    if selected_id:
        selected_results = [
            result
            for result in results
            if result["document_id"] == selected_id
        ]

        other_results = [
            result
            for result in results
            if result["document_id"] != selected_id
        ]

        if selected_results:
            results = selected_results + other_results

    best = results[0]

    direct_answer = (
        st.session_state.get("highlight_target")
        or best.get("direct_sentence")
        or extract_direct_sentence(
            best["text"],
            active_query,
        )
    )

    st.markdown(
        f"""
        <div class="exact-answer">
            <div style="display:flex;justify-content:space-between;
                        align-items:center;">
                <div class="answer-label">
                    Direct Extracted Answer
                </div>
                <span class="answer-pill">
                    Top Confidence Match
                </span>
            </div>

            <div class="answer-text">
                "{html.escape(direct_answer)}"
            </div>

            <div class="answer-meta">
                Source:
                <b>{html.escape(best["filename"])}</b>
                &nbsp;•&nbsp;
                Page {best["page_number"]}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns(
        [1, 1.25],
        gap="large",
    )

    with left:
        st.markdown(
            f"**Supporting Citations** "
            f"({len(results)} matches)"
        )

        for index, result in enumerate(results):
            is_top = index == 0

            label = (
                "TOP MATCH"
                if is_top
                else "SEARCH RESULT"
            )

            label_color = (
                "#00a982"
                if is_top
                else "#0369a1"
            )

            with st.container(
                border=True,
            ):
                st.markdown(
                    f"""
                    <div class="result-label"
                         style="color:{label_color};">
                        {label}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                if st.button(
                    f"📄 {result['filename']}",
                    key=(
                        f"citation_{result['id']}_"
                        f"{index}_{hash(active_query)}"
                    ),
                    use_container_width=True,
                ):
                    st.session_state.selected_document = (
                        result["document_id"]
                    )
                    st.session_state.viewer_page = (
                        result["page_number"]
                    )
                    st.session_state.highlight_target = (
                        result.get("direct_sentence")
                        or result["text"]
                    )
                    st.rerun()

                st.markdown(
                    f"""
                    <div style="font-size:10.5px;
                                color:#64748b;">
                        Page {result["page_number"]}
                        &nbsp;•&nbsp;
                        Match {result["raw_score"]:.1%}
                    </div>

                    <div class="result-snippet">
                        {html.escape(result["snippet"])}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    with right:
        current_page = (
            st.session_state.get("viewer_page")
            or best["page_number"]
        )

        current_doc = get_document(
            best["document_id"]
        )

        if not current_doc:
            st.error("The source document is unavailable.")
            return

        max_pages = current_doc["page_count"]

        current_page = max(
            1,
            min(current_page, max_pages),
        )

        st.markdown(
            f"""
            <div class="reader-toolbar">
                <div class="reader-title">
                    📄 {html.escape(current_doc["filename"])}
                </div>
                <div class="reader-meta">
                    Page {current_page} of {max_pages}
                    &nbsp;•&nbsp;
                    Matching text is highlighted
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        nav1, nav2, nav3, zoom_col = st.columns(
            [1, 1.4, 1, 1.4],
        )

        with nav1:
            if st.button(
                "◀ Prev",
                disabled=current_page <= 1,
                use_container_width=True,
            ):
                st.session_state.viewer_page = (
                    current_page - 1
                )
                st.session_state.highlight_target = ""
                st.rerun()

        with nav2:
            st.markdown(
                f"""
                <div style="text-align:center;
                            padding-top:8px;
                            font-size:11px;
                            color:#475569;">
                    Page <b>{current_page}</b> / {max_pages}
                </div>
                """,
                unsafe_allow_html=True,
            )

        with nav3:
            if st.button(
                "Next ▶",
                disabled=current_page >= max_pages,
                use_container_width=True,
            ):
                st.session_state.viewer_page = (
                    current_page + 1
                )
                st.session_state.highlight_target = ""
                st.rerun()

        with zoom_col:
            zoom = st.selectbox(
                "Zoom",
                [100, 125, 150, 175, 200],
                index=2,
                format_func=lambda value: f"{value}%",
                label_visibility="collapsed",
            )

        image = render_pdf_page(
            current_doc["stored_path"],
            current_page,
            st.session_state.get(
                "highlight_target",
                direct_answer,
            ),
            zoom,
        )

        if image:
            st.image(
                image,
                use_container_width=True,
            )
            st.caption(
                "🟢 Matching text is automatically highlighted."
            )
        else:
            st.error(
                "Unable to render this PDF page."
            )

        pdf_path = Path(
            current_doc["stored_path"]
        )

        if pdf_path.exists():
            st.download_button(
                "⬇ Download Source PDF",
                data=pdf_path.read_bytes(),
                file_name=current_doc["filename"],
                mime="application/pdf",
                use_container_width=True,
            )


# ============================================================
# DOCUMENT DETAIL
# ============================================================

def document_page() -> None:
    document_id = st.session_state.get(
        "selected_document",
    )

    if not document_id:
        st.session_state.page = "Home"
        st.rerun()

    document = get_document(document_id)

    if not document:
        st.error("Document not found.")
        return

    st.markdown(
        f"""
        <div class="section-title">
            📄 {html.escape(document["title"])}
        </div>
        <div style="color:#6b8290;font-size:11px;margin-bottom:10px;">
            {html.escape(document["category"])}
            &nbsp;•&nbsp;
            v{html.escape(document["version"])}
            &nbsp;•&nbsp;
            {document["page_count"]} pages
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("← Back to Knowledge Base"):
        st.session_state.page = "Home"
        st.session_state.selected_document = None
        st.rerun()

    left, right = st.columns(
        [1.6, 1],
        gap="large",
    )

    with left:
        page = st.session_state.get(
            "viewer_page",
            1,
        )

        page = max(
            1,
            min(page, document["page_count"]),
        )

        nav1, nav2, nav3, zoom_col = st.columns(
            [1, 1.3, 1, 1.3]
        )

        with nav1:
            if st.button(
                "◀ Prev",
                disabled=page <= 1,
                use_container_width=True,
            ):
                st.session_state.viewer_page = page - 1
                st.rerun()

        with nav2:
            st.markdown(
                f"<div style='text-align:center;padding-top:8px;'>"
                f"Page <b>{page}</b> / {document['page_count']}"
                f"</div>",
                unsafe_allow_html=True,
            )

        with nav3:
            if st.button(
                "Next ▶",
                disabled=page >= document["page_count"],
                use_container_width=True,
            ):
                st.session_state.viewer_page = page + 1
                st.rerun()

        with zoom_col:
            zoom = st.selectbox(
                "Zoom",
                [100, 125, 150, 175, 200],
                index=2,
                key="document_zoom",
                format_func=lambda value: f"{value}%",
                label_visibility="collapsed",
            )

        image = render_pdf_page(
            document["stored_path"],
            page,
            st.session_state.get(
                "highlight_target",
                "",
            ),
            zoom,
        )

        if image:
            st.image(
                image,
                use_container_width=True,
            )

    with right:
        st.markdown("### Document Information")

        st.markdown(
            f"""
            <div class="doc-card">
                <b>Category</b><br>
                {html.escape(document["category"])}
                <br><br>

                <b>Version</b><br>
                {html.escape(document["version"])}
                <br><br>

                <b>Pages</b><br>
                {document["page_count"]}
                <br><br>

                <b>Tags</b><br>
                {html.escape(document["tags"] or "None")}
                <br><br>

                <b>Description</b><br>
                {html.escape(
                    document["description"]
                    or "No description provided."
                )}
            </div>
            """,
            unsafe_allow_html=True,
        )

        path = Path(document["stored_path"])

        if path.exists():
            st.download_button(
                "⬇ Download PDF",
                data=path.read_bytes(),
                file_name=document["filename"],
                mime="application/pdf",
                use_container_width=True,
            )


# ============================================================
# ADMIN LOGIN
# ============================================================

def admin_login_page() -> None:
    st.markdown("### Administrator Verification")
    st.caption(
        "Enter the management PIN to manage knowledge-base documents."
    )

    with st.form("admin_pin_form"):
        pin = st.text_input(
            "Master PIN",
            type="password",
            placeholder="Enter admin PIN",
        )

        verify, cancel = st.columns(2)

        with verify:
            submitted = st.form_submit_button(
                "Verify",
                type="primary",
                use_container_width=True,
            )

        with cancel:
            cancelled = st.form_submit_button(
                "Cancel",
                use_container_width=True,
            )

    if submitted:
        if pin == ADMIN_PIN:
            st.session_state.admin_authenticated = True
            st.session_state.page = "Manage Documents"
            st.rerun()
        else:
            st.error("Access denied: Invalid PIN.")

    if cancelled:
        st.session_state.page = "Home"
        st.rerun()


# ============================================================
# ADMIN DOCUMENT CENTER
# ============================================================

def manage_documents_page() -> None:
    if not admin_is_authorized():
        st.session_state.page = "Admin Login"
        st.rerun()

    st.markdown(
        """
        <div class="admin-banner">
            <strong>Administrator Mode</strong><br>
            Upload, replace, categorize, version, and remove
            knowledge-base PDF sources.
        </div>
        """,
        unsafe_allow_html=True,
    )

    header_col, back_col = st.columns(
        [4, 1],
    )

    with header_col:
        st.markdown("### Knowledge Repository Management")
        st.caption(
            "Every PDF is indexed page-by-page for search and source navigation."
        )

    with back_col:
        if st.button(
            "← Search",
            use_container_width=True,
        ):
            st.session_state.page = "Home"
            st.rerun()

    upload_col, list_col = st.columns(
        [1, 1.4],
        gap="large",
    )

    with upload_col:
        st.markdown("#### ☁️ Add Document")

        category = st.selectbox(
            "Category",
            CATEGORIES,
            key="admin_upload_category",
        )

        files = st.file_uploader(
            "Upload PDF Documents",
            type=["pdf"],
            accept_multiple_files=True,
            key="admin_upload_files",
        )

        if files:
            st.caption(
                f"{len(files)} PDF(s) selected."
            )

        if st.button(
            "Index Documents",
            type="primary",
            use_container_width=True,
            disabled=not files,
        ):
            progress = st.progress(0)

            for index, file in enumerate(files):
                if file.size > MAX_UPLOAD_MB * 1024 * 1024:
                    st.warning(
                        f"{file.name}: exceeds {MAX_UPLOAD_MB} MB."
                    )
                    progress.progress(
                        (index + 1) / len(files)
                    )
                    continue

                pdf_bytes = file.getvalue()

                default_title = Path(
                    file.name
                ).stem.replace("_", " ").strip()

                ok, message = add_document(
                    filename=file.name,
                    pdf_bytes=pdf_bytes,
                    category=category,
                    title=default_title,
                    description="",
                    tags="",
                    version="1.0",
                    uploaded_by="Administrator",
                )

                if ok:
                    st.success(message)
                else:
                    st.warning(message)

                progress.progress(
                    (index + 1) / len(files)
                )

            clear_search_cache()
            st.rerun()

    with list_col:
        st.markdown("#### 📚 Indexed Documents")

        docs = get_documents()

        if not docs:
            st.info(
                "No documents are currently indexed."
            )
            return

        admin_filter = st.text_input(
            "Filter documents",
            placeholder="Title, filename, category, tag...",
        )

        if admin_filter.strip():
            query_lower = admin_filter.lower()

            docs = [
                document
                for document in docs
                if query_lower
                in (
                    document["title"]
                    + " "
                    + document["filename"]
                    + " "
                    + document["category"]
                    + " "
                    + document["tags"]
                ).lower()
            ]

        for document in docs:
            with st.expander(
                f"📄 {document['title']}  •  "
                f"{document['category']}  •  "
                f"v{document['version']}"
            ):
                title = st.text_input(
                    "Title",
                    value=document["title"],
                    key=f"title_{document['id']}",
                )

                category = st.selectbox(
                    "Category",
                    CATEGORIES,
                    index=(
                        CATEGORIES.index(
                            document["category"]
                        )
                        if document["category"] in CATEGORIES
                        else 0
                    ),
                    key=f"category_{document['id']}",
                )

                version = st.text_input(
                    "Version",
                    value=document["version"],
                    key=f"version_{document['id']}",
                )

                description = st.text_area(
                    "Description",
                    value=document["description"],
                    key=f"description_{document['id']}",
                )

                tags = st.text_input(
                    "Tags",
                    value=document["tags"],
                    key=f"tags_{document['id']}",
                )

                replacement = st.file_uploader(
                    "Replace PDF (optional)",
                    type=["pdf"],
                    key=f"replacement_{document['id']}",
                )

                save_col, download_col, delete_col = st.columns(3)

                with save_col:
                    if st.button(
                        "Save Changes",
                        type="primary",
                        key=f"save_{document['id']}",
                        use_container_width=True,
                    ):
                        if replacement:
                            replacement_bytes = (
                                replacement.getvalue()
                            )

                            ok, message = replace_document(
                                document_id=document["id"],
                                pdf_bytes=replacement_bytes,
                                filename=replacement.name,
                                category=category,
                                title=title,
                                description=description,
                                tags=tags,
                                version=version,
                                uploaded_by="Administrator",
                            )
                        else:
                            conn = db()

                            conn.execute(
                                """
                                UPDATE documents
                                SET title = ?,
                                    category = ?,
                                    description = ?,
                                    tags = ?,
                                    version = ?,
                                    updated_at = ?,
                                    uploaded_by = ?
                                WHERE id = ?
                                """,
                                (
                                    title,
                                    category,
                                    description,
                                    tags,
                                    version,
                                    now_iso(),
                                    "Administrator",
                                    document["id"],
                                ),
                            )

                            conn.commit()
                            conn.close()

                            record_event(
                                "metadata_update",
                                document["id"],
                                document["filename"],
                            )

                            ok = True
                            message = "Document metadata updated."

                        if ok:
                            clear_search_cache()
                            st.success(message)
                            st.rerun()
                        else:
                            st.error(message)

                with download_col:
                    path = Path(
                        document["stored_path"]
                    )

                    if path.exists():
                        st.download_button(
                            "Download",
                            data=path.read_bytes(),
                            file_name=document["filename"],
                            mime="application/pdf",
                            key=f"download_{document['id']}",
                            use_container_width=True,
                        )

                with delete_col:
                    if st.button(
                        "Delete",
                        key=f"delete_{document['id']}",
                        use_container_width=True,
                    ):
                        st.session_state[
                            f"confirm_delete_{document['id']}"
                        ] = True

                if st.session_state.get(
                    f"confirm_delete_{document['id']}",
                    False,
                ):
                    st.warning(
                        "This removes the document and all of its search chunks."
                    )

                    yes, no = st.columns(2)

                    with yes:
                        if st.button(
                            "Confirm Delete",
                            type="primary",
                            key=f"confirm_yes_{document['id']}",
                            use_container_width=True,
                        ):
                            delete_document(
                                document["id"]
                            )

                            st.session_state.pop(
                                f"confirm_delete_{document['id']}",
                                None,
                            )

                            clear_search_cache()
                            st.rerun()

                    with no:
                        if st.button(
                            "Cancel",
                            key=f"confirm_no_{document['id']}",
                            use_container_width=True,
                        ):
                            st.session_state.pop(
                                f"confirm_delete_{document['id']}",
                                None,
                            )
                            st.rerun()


# ============================================================
# ADMIN ANALYTICS
# ============================================================

def analytics_page() -> None:
    if not admin_is_authorized():
        st.session_state.page = "Admin Login"
        st.rerun()

    docs = get_documents()

    st.markdown(
        '<div class="section-title">📊 Knowledge Base Analytics</div>',
        unsafe_allow_html=True,
    )

    if not docs:
        st.info("No data available yet.")
        return

    col1, col2 = st.columns(2)

    with col1:
        category_data = (
            pd.DataFrame(docs)
            .groupby("category")
            .size()
            .reset_index(name="Documents")
            .set_index("category")
        )

        st.markdown("#### Documents by Category")
        st.bar_chart(category_data)

    with col2:
        view_data = (
            pd.DataFrame(
                [
                    {
                        "Document": document["title"],
                        "Views": document["view_count"],
                    }
                    for document in docs
                ]
            )
            .sort_values(
                "Views",
                ascending=False,
            )
            .head(10)
            .set_index("Document")
        )

        st.markdown("#### Most Viewed")
        st.bar_chart(view_data)

    st.markdown("#### Repository Inventory")

    inventory = pd.DataFrame(
        [
            {
                "Document": document["title"],
                "Category": document["category"],
                "Version": document["version"],
                "Pages": document["page_count"],
                "Views": document["view_count"],
                "Updated": document["updated_at"][:10],
            }
            for document in docs
        ]
    )

    st.dataframe(
        inventory,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# ROUTER
# ============================================================

render_topbar()

page = st.session_state.get(
    "page",
    "Home",
)

if page == "Home":
    home_page()

elif page == "Search":
    search_page()

elif page == "Document":
    document_page()

elif page == "Admin Login":
    admin_login_page()

elif page == "Manage Documents":
    manage_documents_page()

elif page == "Analytics":
    analytics_page()

else:
    home_page()
