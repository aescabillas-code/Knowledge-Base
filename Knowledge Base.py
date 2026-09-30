"""
HPE Knowledge Base — Streamlit single-file application

Features
--------
- Knowledge Base dashboard styled after the supplied UI mockup.
- PDF upload and text extraction.
- Search across extracted PDF text and document metadata.
- Interactive category cards and document cards.
- Built-in PDF preview with browser/Streamlit PDF rendering fallback.
- One-time access codes for standard users:
    * Each code can be used once.
    * Codes expire automatically.
    * Codes are stored as hashes, not plain text.
- Admin login:
    * Admin can upload new PDFs.
    * Admin can replace/update an existing PDF.
    * Admin can edit title/category/description/tags.
    * Admin can delete documents.
    * Admin can generate/revoke one-time access codes.
- Persistent SQLite database and local PDF storage.
- Optional MongoDB support can be added later without changing the UI.

Install
-------
pip install streamlit pymupdf bcrypt pandas

Run
---
streamlit run "Knowledge Base.py"

Optional Streamlit secrets
--------------------------
[admin]
email = "admin@example.com"
password = "ChangeThisPassword"

[app]
access_code_ttl_hours = 24
max_upload_mb = 50

For Streamlit Cloud, add these values under Settings > Secrets.

IMPORTANT
---------
The local SQLite/PDF directory is persistent on a normal server/VM.
Streamlit Community Cloud has ephemeral local storage, so use an external
database/object store for production persistence.
"""

from __future__ import annotations

import base64
import hashlib
import html
import os
import secrets
import sqlite3
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Optional

import bcrypt
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# ============================================================
# APP CONFIG
# ============================================================

st.set_page_config(
    page_title="HPE Knowledge Base",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_NAME = "Knowledge Base"
DATA_DIR = Path(os.getenv("KB_DATA_DIR", "kb_data"))
PDF_DIR = DATA_DIR / "pdfs"
DB_PATH = DATA_DIR / "knowledge_base.db"

DATA_DIR.mkdir(parents=True, exist_ok=True)
PDF_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_CATEGORIES = [
    "Account Management",
    "Licensing",
    "Portal & Access",
    "Technical Support",
    "HPE GreenLake",
    "Product Guides",
    "Policies & Procedures",
    "Troubleshooting",
    "Other",
]

CATEGORY_ICONS = {
    "Account Management": "👤",
    "Licensing": "🔑",
    "Portal & Access": "🖥️",
    "Technical Support": "🛠️",
    "HPE GreenLake": "☁️",
    "Product Guides": "📖",
    "Policies & Procedures": "🛡️",
    "Troubleshooting": "⚙️",
    "Other": "📄",
}


# ============================================================
# SECRETS / ADMIN CONFIG
# ============================================================

def get_secret(path: tuple[str, ...], default: str = "") -> str:
    try:
        value: Any = st.secrets
        for key in path:
            value = value[key]
        return str(value)
    except Exception:
        return default


ADMIN_EMAIL = get_secret(
    ("admin", "email"),
    os.getenv("KB_ADMIN_EMAIL", "admin@example.com"),
)

ADMIN_PASSWORD = get_secret(
    ("admin", "password"),
    os.getenv("KB_ADMIN_PASSWORD", "ChangeThisPassword"),
)

ACCESS_CODE_TTL_HOURS = int(
    get_secret(
        ("app", "access_code_ttl_hours"),
        os.getenv("KB_ACCESS_CODE_TTL_HOURS", "24"),
    )
)

MAX_UPLOAD_MB = int(
    get_secret(
        ("app", "max_upload_mb"),
        os.getenv("KB_MAX_UPLOAD_MB", "50"),
    )
)


# ============================================================
# DATABASE
# ============================================================

def db() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH, check_same_thread=False)
    connection.row_factory = sqlite3.Row
    return connection


def init_db() -> None:
    connection = db()
    cursor = connection.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            filename TEXT NOT NULL,
            stored_filename TEXT NOT NULL UNIQUE,
            category TEXT NOT NULL,
            description TEXT DEFAULT '',
            tags TEXT DEFAULT '',
            extracted_text TEXT DEFAULT '',
            file_size INTEGER DEFAULT 0,
            page_count INTEGER DEFAULT 0,
            version TEXT DEFAULT '1.0',
            uploaded_by TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            view_count INTEGER DEFAULT 0,
            is_active INTEGER DEFAULT 1
        )
        """
    )

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS access_codes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code_hash TEXT NOT NULL,
            created_by TEXT NOT NULL,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            used_at TEXT,
            revoked INTEGER DEFAULT 0
        )
        """
    )

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS document_views (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER NOT NULL,
            viewed_at TEXT NOT NULL
        )
        """
    )

    connection.commit()
    connection.close()


init_db()


# ============================================================
# SECURITY HELPERS
# ============================================================

def now_utc() -> datetime:
    return datetime.now(timezone.utc)


def iso_now() -> str:
    return now_utc().isoformat()


def hash_access_code(code: str) -> str:
    return hashlib.sha256(code.strip().encode("utf-8")).hexdigest()


def verify_access_code(code: str, stored_hash: str) -> bool:
    return secrets.compare_digest(hash_access_code(code), stored_hash)


def password_matches(password: str) -> bool:
    return secrets.compare_digest(
        password,
        ADMIN_PASSWORD,
    )


def generate_access_code() -> str:
    # Easy-to-type format: KB-XXXX-XXXX
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    a = "".join(secrets.choice(alphabet) for _ in range(4))
    b = "".join(secrets.choice(alphabet) for _ in range(4))
    return f"KB-{a}-{b}"


def create_access_code(admin_email: str) -> str:
    code = generate_access_code()
    expires = now_utc() + timedelta(hours=ACCESS_CODE_TTL_HOURS)

    connection = db()
    connection.execute(
        """
        INSERT INTO access_codes
        (code_hash, created_by, created_at, expires_at)
        VALUES (?, ?, ?, ?)
        """,
        (
            hash_access_code(code),
            admin_email,
            iso_now(),
            expires.isoformat(),
        ),
    )
    connection.commit()
    connection.close()
    return code


def consume_access_code(code: str) -> bool:
    code = code.strip()
    if not code:
        return False

    connection = db()
    rows = connection.execute(
        """
        SELECT id, code_hash, expires_at, used_at, revoked
        FROM access_codes
        WHERE used_at IS NULL
          AND revoked = 0
        ORDER BY id DESC
        """
    ).fetchall()

    matched_id: Optional[int] = None

    for row in rows:
        if verify_access_code(code, row["code_hash"]):
            try:
                expires_at = datetime.fromisoformat(row["expires_at"])
                if expires_at.tzinfo is None:
                    expires_at = expires_at.replace(tzinfo=timezone.utc)

                if expires_at <= now_utc():
                    connection.close()
                    return False
            except Exception:
                connection.close()
                return False

            matched_id = row["id"]
            break

    if matched_id is None:
        connection.close()
        return False

    # Atomic-ish single-use consumption inside one write.
    result = connection.execute(
        """
        UPDATE access_codes
        SET used_at = ?
        WHERE id = ?
          AND used_at IS NULL
          AND revoked = 0
        """,
        (iso_now(), matched_id),
    )

    connection.commit()
    connection.close()

    return result.rowcount == 1


def cleanup_expired_codes() -> None:
    connection = db()
    connection.execute(
        """
        UPDATE access_codes
        SET revoked = 1
        WHERE used_at IS NULL
          AND revoked = 0
          AND expires_at <= ?
        """,
        (iso_now(),),
    )
    connection.commit()
    connection.close()


cleanup_expired_codes()


# ============================================================
# PDF HELPERS
# ============================================================

def safe_filename(name: str) -> str:
    name = Path(name).name
    cleaned = "".join(
        c if c.isalnum() or c in "._- " else "_"
        for c in name
    )
    return cleaned.strip() or "document.pdf"


def extract_pdf(pdf_bytes: bytes) -> tuple[str, int]:
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = len(document)

    chunks = []
    for page in document:
        text = page.get_text("text")
        if text:
            chunks.append(text)

    text = "\n".join(chunks)
    document.close()

    return text, pages


def save_pdf(pdf_bytes: bytes, stored_filename: str) -> Path:
    path = PDF_DIR / stored_filename
    path.write_bytes(pdf_bytes)
    return path


def load_pdf(stored_filename: str) -> Optional[bytes]:
    path = PDF_DIR / stored_filename
    if not path.exists():
        return None
    return path.read_bytes()


def make_storage_name(original_filename: str) -> str:
    stem = Path(original_filename).stem
    ext = Path(original_filename).suffix.lower() or ".pdf"
    unique = secrets.token_hex(8)
    return f"{safe_filename(stem)}_{unique}{ext}"


def document_preview(pdf_bytes: bytes) -> None:
    """
    Render a browser PDF viewer where supported.
    Falls back to first-page image if the embedded PDF viewer is blocked.
    """
    if hasattr(st, "pdf"):
        try:
            st.pdf(pdf_bytes, height=680)
            return
        except Exception:
            pass

    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = document[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(1.4, 1.4), alpha=False)
        image_bytes = pix.tobytes("png")
        document.close()
        st.image(image_bytes, use_container_width=True)
        st.caption("PDF preview fallback — use Download PDF to view the full document.")
    except Exception as exc:
        st.warning(f"Unable to render preview: {exc}")


# ============================================================
# DOCUMENT DATABASE OPERATIONS
# ============================================================

def add_document(
    *,
    title: str,
    filename: str,
    stored_filename: str,
    category: str,
    description: str,
    tags: str,
    extracted_text: str,
    file_size: int,
    page_count: int,
    version: str,
    uploaded_by: str,
) -> int:
    timestamp = iso_now()

    connection = db()
    cursor = connection.execute(
        """
        INSERT INTO documents (
            title, filename, stored_filename, category, description,
            tags, extracted_text, file_size, page_count, version,
            uploaded_by, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            title,
            filename,
            stored_filename,
            category,
            description,
            tags,
            extracted_text,
            file_size,
            page_count,
            version,
            uploaded_by,
            timestamp,
            timestamp,
        ),
    )

    connection.commit()
    document_id = cursor.lastrowid
    connection.close()
    return int(document_id)


def update_document(
    document_id: int,
    *,
    title: str,
    category: str,
    description: str,
    tags: str,
    version: str,
    uploaded_by: str,
    pdf_bytes: Optional[bytes] = None,
) -> None:
    connection = db()

    if pdf_bytes is not None:
        text, pages = extract_pdf(pdf_bytes)

        row = connection.execute(
            "SELECT stored_filename FROM documents WHERE id = ?",
            (document_id,),
        ).fetchone()

        if not row:
            connection.close()
            raise ValueError("Document no longer exists.")

        stored_filename = row["stored_filename"]
        save_pdf(pdf_bytes, stored_filename)

        connection.execute(
            """
            UPDATE documents
            SET title = ?,
                category = ?,
                description = ?,
                tags = ?,
                version = ?,
                extracted_text = ?,
                file_size = ?,
                page_count = ?,
                uploaded_by = ?,
                updated_at = ?
            WHERE id = ?
            """,
            (
                title,
                category,
                description,
                tags,
                version,
                text,
                len(pdf_bytes),
                pages,
                uploaded_by,
                iso_now(),
                document_id,
            ),
        )
    else:
        connection.execute(
            """
            UPDATE documents
            SET title = ?,
                category = ?,
                description = ?,
                tags = ?,
                version = ?,
                uploaded_by = ?,
                updated_at = ?
            WHERE id = ?
            """,
            (
                title,
                category,
                description,
                tags,
                version,
                uploaded_by,
                iso_now(),
                document_id,
            ),
        )

    connection.commit()
    connection.close()


def delete_document(document_id: int) -> None:
    connection = db()

    row = connection.execute(
        "SELECT stored_filename FROM documents WHERE id = ?",
        (document_id,),
    ).fetchone()

    if row:
        try:
            (PDF_DIR / row["stored_filename"]).unlink(missing_ok=True)
        except Exception:
            pass

    connection.execute(
        "UPDATE documents SET is_active = 0 WHERE id = ?",
        (document_id,),
    )

    connection.commit()
    connection.close()


def get_documents() -> list[dict[str, Any]]:
    connection = db()
    rows = connection.execute(
        """
        SELECT *
        FROM documents
        WHERE is_active = 1
        ORDER BY updated_at DESC
        """
    ).fetchall()
    connection.close()
    return [dict(row) for row in rows]


def get_document(document_id: int) -> Optional[dict[str, Any]]:
    connection = db()
    row = connection.execute(
        """
        SELECT *
        FROM documents
        WHERE id = ? AND is_active = 1
        """,
        (document_id,),
    ).fetchone()
    connection.close()

    return dict(row) if row else None


def increment_view(document_id: int) -> None:
    connection = db()
    connection.execute(
        """
        UPDATE documents
        SET view_count = view_count + 1
        WHERE id = ?
        """,
        (document_id,),
    )
    connection.execute(
        """
        INSERT INTO document_views(document_id, viewed_at)
        VALUES (?, ?)
        """,
        (document_id, iso_now()),
    )
    connection.commit()
    connection.close()


def search_documents(query: str) -> list[dict[str, Any]]:
    docs = get_documents()
    if not query.strip():
        return docs

    q = query.lower().strip()
    terms = [term for term in q.split() if term]

    scored = []

    for doc in docs:
        haystack = " ".join(
            [
                doc.get("title", ""),
                doc.get("category", ""),
                doc.get("description", ""),
                doc.get("tags", ""),
                doc.get("extracted_text", ""),
            ]
        ).lower()

        score = 0

        for term in terms:
            if term in doc.get("title", "").lower():
                score += 20
            if term in doc.get("category", "").lower():
                score += 10
            if term in doc.get("tags", "").lower():
                score += 8
            if term in doc.get("description", "").lower():
                score += 5
            if term in doc.get("extracted_text", "").lower():
                score += 2

        if score:
            scored.append((score, doc))

    scored.sort(
        key=lambda item: (
            -item[0],
            item[1].get("updated_at", ""),
        )
    )

    return [doc for _, doc in scored]


# ============================================================
# UI CSS
# ============================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        /* ---------- GLOBAL ---------- */
        .stApp {
            background: #f5f8fa;
        }

        header[data-testid="stHeader"] {
            background: rgba(255,255,255,0.96);
        }

        [data-testid="stSidebar"] {
            background: #06383d;
            border-right: 0;
        }

        [data-testid="stSidebar"] * {
            color: #f5ffff !important;
        }

        [data-testid="stSidebar"] .stButton > button {
            background: transparent;
            border: 0;
            text-align: left;
            width: 100%;
            border-radius: 10px;
        }

        [data-testid="stSidebar"] .stButton > button:hover {
            background: rgba(255,255,255,0.10);
        }

        /* Hide Streamlit branding/footer. */
        footer {
            visibility: hidden;
        }

        /* ---------- TOP BRAND ---------- */
        .kb-brand {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 4px 0 14px 0;
        }

        .kb-logo {
            font-weight: 900;
            font-size: 30px;
            letter-spacing: -2px;
            color: white;
        }

        .kb-logo span {
            color: #01a982;
        }

        .kb-small {
            color: #b8d4d7;
            font-size: 12px;
        }

        /* ---------- PAGE HEADER ---------- */
        .page-title {
            font-size: 34px;
            font-weight: 800;
            color: #102a43;
            margin-bottom: 0;
        }

        .page-subtitle {
            color: #627d98;
            margin-top: 2px;
            margin-bottom: 16px;
        }

        /* ---------- HERO ---------- */
        .hero {
            background:
                radial-gradient(circle at 85% 20%, rgba(1,169,130,.30), transparent 30%),
                linear-gradient(115deg, #073b40, #006b70 58%, #008f83);
            border-radius: 16px;
            padding: 32px 36px 28px 36px;
            color: white;
            margin-bottom: 18px;
            box-shadow: 0 12px 35px rgba(3,63,67,.15);
        }

        .hero h1 {
            font-size: 34px;
            line-height: 1.15;
            margin: 0 0 8px 0;
            color: white;
        }

        .hero p {
            color: #d7f1f1;
            margin: 0 0 22px 0;
            font-size: 16px;
        }

        .hero-chip {
            display: inline-block;
            padding: 6px 12px;
            margin: 6px 5px 0 0;
            background: rgba(255,255,255,.13);
            border: 1px solid rgba(255,255,255,.16);
            border-radius: 20px;
            color: white;
            font-size: 12px;
        }

        /* ---------- METRIC CARDS ---------- */
        .metric-card {
            background: white;
            border: 1px solid #e3edf2;
            border-radius: 14px;
            padding: 18px;
            min-height: 110px;
            box-shadow: 0 4px 16px rgba(16,42,67,.05);
        }

        .metric-icon {
            width: 42px;
            height: 42px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border-radius: 12px;
            background: #e8f8f4;
            font-size: 22px;
            margin-bottom: 8px;
        }

        .metric-number {
            font-size: 27px;
            font-weight: 800;
            color: #102a43;
        }

        .metric-label {
            font-size: 13px;
            color: #627d98;
        }

        /* ---------- CATEGORY ---------- */
        .category-card {
            background: white;
            border: 1px solid #e3edf2;
            border-radius: 13px;
            padding: 15px 16px;
            min-height: 86px;
            box-shadow: 0 3px 12px rgba(16,42,67,.035);
        }

        .category-icon {
            float: left;
            margin-right: 12px;
            font-size: 24px;
        }

        .category-title {
            font-weight: 750;
            color: #102a43;
            margin-top: 2px;
        }

        .category-count {
            color: #829ab1;
            font-size: 12px;
        }

        /* ---------- DOCUMENT CARD ---------- */
        .doc-card {
            background: white;
            border: 1px solid #e3edf2;
            border-radius: 13px;
            padding: 15px;
            margin-bottom: 10px;
            box-shadow: 0 3px 12px rgba(16,42,67,.035);
        }

        .pdf-badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 40px;
            height: 40px;
            background: #fff0ef;
            color: #d64545;
            border-radius: 10px;
            font-weight: 800;
            font-size: 12px;
            margin-right: 10px;
        }

        .doc-title {
            color: #102a43;
            font-weight: 750;
            font-size: 15px;
        }

        .doc-meta {
            color: #829ab1;
            font-size: 12px;
            margin-top: 3px;
        }

        .section-title {
            color: #102a43;
            font-size: 20px;
            font-weight: 800;
            margin: 14px 0 10px 0;
        }

        /* ---------- ADMIN ---------- */
        .admin-banner {
            background: linear-gradient(100deg,#062f34,#087d77);
            color: white;
            padding: 14px 18px;
            border-radius: 12px;
            margin-bottom: 15px;
        }

        .admin-banner strong {
            color: white;
        }

        /* ---------- LOGIN ---------- */
        .login-wrap {
            max-width: 540px;
            margin: 6vh auto;
        }

        .login-card {
            background: white;
            padding: 35px;
            border-radius: 20px;
            border: 1px solid #e3edf2;
            box-shadow: 0 15px 50px rgba(16,42,67,.10);
        }

        /* ---------- BUTTONS ---------- */
        .stButton > button,
        .stDownloadButton > button {
            border-radius: 9px;
            min-height: 40px;
        }

        /* ---------- PDF ---------- */
        iframe {
            border-radius: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_css()


# ============================================================
# SESSION / AUTH
# ============================================================

def init_session() -> None:
    defaults = {
        "authenticated": False,
        "is_admin": False,
        "auth_method": None,
        "user_email": None,
        "page": "Home",
        "selected_document_id": None,
        "selected_category": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_session()


def sign_out() -> None:
    for key in [
        "authenticated",
        "is_admin",
        "auth_method",
        "user_email",
        "page",
        "selected_document_id",
        "selected_category",
    ]:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()


# ============================================================
# LOGIN
# ============================================================

def login_page() -> None:
    st.markdown(
        """
        <div class="login-wrap">
            <div class="login-card">
                <div style="font-size:42px;font-weight:900;color:#102a43;">
                    H<span style="color:#01a982;">P</span>E
                </div>
                <div style="font-size:30px;font-weight:800;color:#102a43;">
                    Knowledge Base
                </div>
                <div style="color:#627d98;margin:5px 0 25px 0;">
                    Search approved support documentation and guides.
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Use a centered two-column choice without exposing admin credentials.
    left, right = st.columns(2)

    with left:
        st.markdown("### 🔐 One-Time Access")
        st.caption(
            f"Access codes expire after {ACCESS_CODE_TTL_HOURS} hours "
            "and can only be used once."
        )

        code = st.text_input(
            "Access code",
            placeholder="KB-XXXX-XXXX",
            key="user_access_code",
        )

        if st.button(
            "Enter Knowledge Base",
            type="primary",
            use_container_width=True,
        ):
            if consume_access_code(code):
                st.session_state.authenticated = True
                st.session_state.is_admin = False
                st.session_state.auth_method = "one_time_code"
                st.session_state.user_email = "Knowledge Base User"
                st.session_state.page = "Home"
                st.success("Access granted.")
                st.rerun()
            else:
                st.error(
                    "Invalid, expired, revoked, or already-used access code."
                )

    with right:
        st.markdown("### 🛡️ Admin Access")
        st.caption("Administrators can manage and update PDF content.")

        admin_email = st.text_input(
            "Admin email",
            key="admin_email_login",
        )

        admin_password = st.text_input(
            "Admin password",
            type="password",
            key="admin_password_login",
        )

        if st.button(
            "Admin Sign In",
            use_container_width=True,
        ):
            if (
                admin_email.strip().lower() == ADMIN_EMAIL.strip().lower()
                and password_matches(admin_password)
            ):
                st.session_state.authenticated = True
                st.session_state.is_admin = True
                st.session_state.auth_method = "admin"
                st.session_state.user_email = ADMIN_EMAIL
                st.session_state.page = "Home"
                st.success("Admin access granted.")
                st.rerun()
            else:
                st.error("Invalid administrator credentials.")

    st.divider()

    st.info(
        "Need access? Ask an administrator to generate a one-time access "
        "code. Standard users cannot upload or modify documents."
    )


# ============================================================
# SIDEBAR
# ============================================================

def sidebar() -> None:
    with st.sidebar:
        st.markdown(
            """
            <div class="kb-brand">
                <div class="kb-logo">HP<span>E</span></div>
                <div>
                    <div style="font-weight:700;">Knowledge Base</div>
                    <div class="kb-small">Document Intelligence Portal</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("---")

        nav_items = [
            ("🏠", "Home"),
            ("📚", "Browse All"),
            ("▦", "Categories"),
            ("⭐", "Favorites"),
            ("🕘", "Recent"),
        ]

        for icon, label in nav_items:
            if st.button(
                f"{icon}  {label}",
                key=f"nav_{label}",
                use_container_width=True,
            ):
                st.session_state.page = label
                st.session_state.selected_document_id = None
                st.session_state.selected_category = None
                st.rerun()

        if st.session_state.is_admin:
            st.markdown("---")
            st.caption("ADMINISTRATION")

            admin_items = [
                ("☁️", "Upload PDF"),
                ("🗂️", "Manage Content"),
                ("🔐", "Access Codes"),
                ("📊", "Analytics"),
            ]

            for icon, label in admin_items:
                if st.button(
                    f"{icon}  {label}",
                    key=f"admin_nav_{label}",
                    use_container_width=True,
                ):
                    st.session_state.page = label
                    st.session_state.selected_document_id = None
                    st.rerun()

        st.markdown("---")

        if st.button(
            "↪  Sign Out",
            key="sidebar_signout",
            use_container_width=True,
        ):
            sign_out()

        st.markdown(
            """
            <div style="margin-top:25vh;color:#a9c6c9;font-size:11px;">
                HPE Knowledge Base<br>
                PDF-powered support documentation
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# COMMON UI
# ============================================================

def top_header() -> None:
    left, middle, right = st.columns([2.2, 5, 1.4])

    with left:
        st.markdown(
            '<div class="page-title">Knowledge Base</div>'
            '<div class="page-subtitle">Find answers from approved documentation.</div>',
            unsafe_allow_html=True,
        )

    with middle:
        st.write("")
        search = st.text_input(
            "Search",
            placeholder="Search for topics, keywords, or questions...",
            label_visibility="collapsed",
            key="global_search",
        )

        if search.strip():
            st.session_state.search_query = search

    with right:
        st.write("")
        if st.session_state.is_admin:
            st.success("ADMIN")
        else:
            st.info("USER")

    if "search_query" in st.session_state and st.session_state.search_query:
        results = search_documents(st.session_state.search_query)
        st.session_state.search_results = results


def hero() -> None:
    st.markdown(
        """
        <div class="hero">
            <h1>Find the answers you need</h1>
            <p>
                Search the knowledge base, explore topics, or browse by category.
            </p>
        """,
        unsafe_allow_html=True,
    )

    hero_search = st.text_input(
        "Search knowledge base",
        placeholder="Search for solutions, guides, keywords, or questions...",
        label_visibility="collapsed",
        key="hero_search",
    )

    if hero_search.strip():
        st.session_state.search_query = hero_search
        st.session_state.page = "Browse All"

    st.markdown(
        """
            <div style="margin-top:10px;">
                <span class="hero-chip">Licensing</span>
                <span class="hero-chip">Portal Access</span>
                <span class="hero-chip">Account Setup</span>
                <span class="hero-chip">Troubleshooting</span>
                <span class="hero-chip">HPE GreenLake</span>
                <span class="hero-chip">Software Support</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_cards(docs: list[dict[str, Any]]) -> None:
    categories = len(set(d["category"] for d in docs)) if docs else 0
    most_viewed = max((d["view_count"] for d in docs), default=0)

    recent_cutoff = now_utc() - timedelta(days=14)
    recent = 0

    for doc in docs:
        try:
            dt = datetime.fromisoformat(doc["updated_at"])
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            if dt >= recent_cutoff:
                recent += 1
        except Exception:
            pass

    metrics = [
        ("📚", len(docs), "Total Documents"),
        ("📁", categories, "Categories"),
        ("⭐", most_viewed, "Most Viewed"),
        ("☁️", recent, "Recently Added"),
    ]

    cols = st.columns(4)

    for col, (icon, value, label) in zip(cols, metrics):
        with col:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-icon">{icon}</div>
                    <div class="metric-number">{value}</div>
                    <div class="metric-label">{label}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def category_grid(
    docs: list[dict[str, Any]],
    interactive: bool = True,
) -> None:
    st.markdown(
        '<div class="section-title">▦ Browse by Category</div>',
        unsafe_allow_html=True,
    )

    counts = {}
    for doc in docs:
        counts[doc["category"]] = counts.get(doc["category"], 0) + 1

    categories = [
        category
        for category in DEFAULT_CATEGORIES
        if counts.get(category, 0) > 0
    ]

    if not categories:
        st.info("No categories are available yet. Upload a PDF to begin.")
        return

    cols = st.columns(4)

    for index, category in enumerate(categories):
        with cols[index % 4]:
            icon = CATEGORY_ICONS.get(category, "📄")
            count = counts.get(category, 0)

            st.markdown(
                f"""
                <div class="category-card">
                    <span class="category-icon">{icon}</span>
                    <div class="category-title">{html.escape(category)}</div>
                    <div class="category-count">{count} documents</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if interactive and st.button(
                f"Open {category}",
                key=f"category_{index}_{category}",
                use_container_width=True,
            ):
                st.session_state.selected_category = category
                st.session_state.page = "Browse All"
                st.rerun()


def document_row(doc: dict[str, Any], key_prefix: str = "") -> None:
    title = html.escape(doc["title"])
    category = html.escape(doc["category"])
    version = html.escape(doc["version"])
    updated = doc["updated_at"][:10]

    st.markdown(
        f"""
        <div class="doc-card">
            <span class="pdf-badge">PDF</span>
            <span class="doc-title">{title}</span>
            <div class="doc-meta">
                {category} &nbsp;•&nbsp; v{version}
                &nbsp;•&nbsp; Updated {updated}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    cols = st.columns([1, 1, 4])

    with cols[0]:
        if st.button(
            "Open",
            key=f"{key_prefix}open_{doc['id']}",
            use_container_width=True,
        ):
            st.session_state.selected_document_id = doc["id"]
            st.session_state.page = "Document"
            increment_view(doc["id"])
            st.rerun()

    with cols[1]:
        pdf_bytes = load_pdf(doc["stored_filename"])
        if pdf_bytes:
            st.download_button(
                "Download",
                data=pdf_bytes,
                file_name=doc["filename"],
                mime="application/pdf",
                key=f"{key_prefix}download_{doc['id']}",
                use_container_width=True,
            )


# ============================================================
# HOME
# ============================================================

def home_page() -> None:
    docs = get_documents()

    top_header()
    hero()
    metric_cards(docs)

    st.markdown("")
    category_grid(docs)

    st.markdown(
        '<div class="section-title">📄 Recent Documents</div>',
        unsafe_allow_html=True,
    )

    recent_docs = docs[:5]

    if not recent_docs:
        st.info(
            "No documents have been uploaded yet. "
            + (
                "Use Upload PDF from the admin menu."
                if st.session_state.is_admin
                else "Ask an administrator to add approved PDFs."
            )
        )
    else:
        for doc in recent_docs:
            document_row(doc, "recent_")


# ============================================================
# BROWSE / SEARCH
# ============================================================

def browse_page() -> None:
    top_header()

    docs = get_documents()

    selected_category = st.session_state.get("selected_category")

    if selected_category:
        docs = [
            d for d in docs
            if d["category"] == selected_category
        ]

        st.markdown(
            f"### {CATEGORY_ICONS.get(selected_category, '📄')} "
            f"{selected_category}"
        )

        if st.button("← All Documents"):
            st.session_state.selected_category = None
            st.rerun()

    query = st.session_state.get("search_query", "")

    if query:
        docs = search_documents(query)
        st.markdown(f"### Search results for “{html.escape(query)}”")
        st.caption(f"{len(docs)} matching documents")

    if not docs:
        st.info("No matching documents found.")
        return

    for doc in docs:
        document_row(doc, "browse_")


# ============================================================
# CATEGORIES
# ============================================================

def categories_page() -> None:
    top_header()
    docs = get_documents()
    category_grid(docs)


# ============================================================
# RECENT / FAVORITES
# ============================================================

def recent_page() -> None:
    top_header()
    docs = get_documents()

    docs = sorted(
        docs,
        key=lambda d: d.get("updated_at", ""),
        reverse=True,
    )[:20]

    st.markdown("### 🕘 Recently Updated")

    for doc in docs:
        document_row(doc, "recentpage_")


def favorites_page() -> None:
    top_header()
    st.markdown("### ⭐ Favorites")
    st.info(
        "Favorites are ready for extension. "
        "For a production version, add a per-user favorites table "
        "or connect this to your existing employee database."
    )


# ============================================================
# DOCUMENT DETAIL
# ============================================================

def document_page() -> None:
    document_id = st.session_state.get("selected_document_id")

    if not document_id:
        st.session_state.page = "Home"
        st.rerun()

    doc = get_document(document_id)

    if not doc:
        st.error("The selected document no longer exists.")
        if st.button("Return Home"):
            st.session_state.page = "Home"
            st.rerun()
        return

    if st.button("← Back to Knowledge Base"):
        st.session_state.page = "Browse All"
        st.session_state.selected_document_id = None
        st.rerun()

    st.markdown(
        f"""
        <div class="section-title">
            📄 {html.escape(doc["title"])}
        </div>
        <div class="doc-meta">
            {html.escape(doc["category"])}
            &nbsp;•&nbsp; v{html.escape(doc["version"])}
            &nbsp;•&nbsp; {doc["page_count"]} pages
            &nbsp;•&nbsp; {doc["view_count"]} views
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.75, 1])

    pdf_bytes = load_pdf(doc["stored_filename"])

    with left:
        st.markdown("### Document Preview")

        if pdf_bytes:
            document_preview(pdf_bytes)

            st.download_button(
                "⬇ Download PDF",
                data=pdf_bytes,
                file_name=doc["filename"],
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.error("PDF file is missing from storage.")

    with right:
        st.markdown("### Document Information")

        st.markdown(
            f"""
            <div class="doc-card">
                <b>Category</b><br>
                {html.escape(doc["category"])}
                <br><br>
                <b>Version</b><br>
                {html.escape(doc["version"])}
                <br><br>
                <b>Pages</b><br>
                {doc["page_count"]}
                <br><br>
                <b>Tags</b><br>
                {html.escape(doc["tags"] or "None")}
                <br><br>
                <b>Description</b><br>
                {html.escape(doc["description"] or "No description provided.")}
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.session_state.is_admin:
            st.divider()

            if st.button(
                "✏️ Edit Document",
                use_container_width=True,
            ):
                st.session_state.edit_document_id = doc["id"]
                st.session_state.page = "Manage Content"
                st.rerun()


# ============================================================
# ADMIN — UPLOAD
# ============================================================

def upload_page() -> None:
    if not st.session_state.is_admin:
        st.error("Administrator access required.")
        return

    top_header()

    st.markdown(
        """
        <div class="admin-banner">
            <strong>Administrator Mode</strong><br>
            Upload approved PDF documentation. Uploaded PDFs are indexed
            automatically for knowledge-base search.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### ☁️ Upload PDF")

    uploaded = st.file_uploader(
        "Select a PDF",
        type=["pdf"],
        accept_multiple_files=False,
    )

    if uploaded is None:
        st.info("Choose a PDF to begin.")
        return

    if uploaded.size > MAX_UPLOAD_MB * 1024 * 1024:
        st.error(
            f"This PDF exceeds the configured {MAX_UPLOAD_MB} MB limit."
        )
        return

    st.success(
        f"Selected: {uploaded.name} "
        f"({uploaded.size / 1024 / 1024:.2f} MB)"
    )

    default_title = Path(uploaded.name).stem.replace("_", " ").strip()

    col1, col2 = st.columns(2)

    with col1:
        title = st.text_input(
            "Document title",
            value=default_title,
        )

        category = st.selectbox(
            "Category",
            DEFAULT_CATEGORIES,
        )

        version = st.text_input(
            "Version",
            value="1.0",
        )

    with col2:
        description = st.text_area(
            "Description",
            placeholder="Short description of what this document contains...",
        )

        tags = st.text_input(
            "Tags",
            placeholder="licensing, portal, troubleshooting",
        )

    if st.button(
        "Upload & Index PDF",
        type="primary",
        use_container_width=True,
    ):
        pdf_bytes = uploaded.getvalue()

        try:
            with st.spinner("Extracting PDF text and indexing document..."):
                extracted_text, page_count = extract_pdf(pdf_bytes)

                stored_filename = make_storage_name(uploaded.name)
                save_pdf(pdf_bytes, stored_filename)

                add_document(
                    title=title.strip() or default_title,
                    filename=uploaded.name,
                    stored_filename=stored_filename,
                    category=category,
                    description=description.strip(),
                    tags=tags.strip(),
                    extracted_text=extracted_text,
                    file_size=len(pdf_bytes),
                    page_count=page_count,
                    version=version.strip() or "1.0",
                    uploaded_by=st.session_state.user_email,
                )

            st.success(
                f"{uploaded.name} uploaded successfully and indexed."
            )
            st.session_state.page = "Manage Content"
            st.rerun()

        except Exception as exc:
            st.error(f"Upload failed: {exc}")


# ============================================================
# ADMIN — MANAGE CONTENT
# ============================================================

def manage_content_page() -> None:
    if not st.session_state.is_admin:
        st.error("Administrator access required.")
        return

    top_header()

    st.markdown(
        '<div class="section-title">🗂️ Manage Content</div>',
        unsafe_allow_html=True,
    )

    docs = get_documents()

    if not docs:
        st.info("No documents have been uploaded.")
        return

    # Search within admin content.
    admin_search = st.text_input(
        "Find document",
        placeholder="Search title, category, tags...",
    )

    filtered = docs

    if admin_search.strip():
        filtered = search_documents(admin_search)

    st.caption(f"{len(filtered)} document(s)")

    for doc in filtered:
        with st.expander(
            f"📄 {doc['title']}  •  {doc['category']}  •  v{doc['version']}"
        ):
            left, right = st.columns(2)

            with left:
                title = st.text_input(
                    "Title",
                    value=doc["title"],
                    key=f"edit_title_{doc['id']}",
                )

                category = st.selectbox(
                    "Category",
                    DEFAULT_CATEGORIES,
                    index=(
                        DEFAULT_CATEGORIES.index(doc["category"])
                        if doc["category"] in DEFAULT_CATEGORIES
                        else len(DEFAULT_CATEGORIES) - 1
                    ),
                    key=f"edit_category_{doc['id']}",
                )

                version = st.text_input(
                    "Version",
                    value=doc["version"],
                    key=f"edit_version_{doc['id']}",
                )

            with right:
                description = st.text_area(
                    "Description",
                    value=doc["description"],
                    key=f"edit_description_{doc['id']}",
                )

                tags = st.text_input(
                    "Tags",
                    value=doc["tags"],
                    key=f"edit_tags_{doc['id']}",
                )

                replacement = st.file_uploader(
                    "Replace PDF (optional)",
                    type=["pdf"],
                    key=f"replace_pdf_{doc['id']}",
                )

            action_col1, action_col2, action_col3 = st.columns(3)

            with action_col1:
                if st.button(
                    "Save Changes",
                    type="primary",
                    key=f"save_{doc['id']}",
                    use_container_width=True,
                ):
                    pdf_bytes = (
                        replacement.getvalue()
                        if replacement is not None
                        else None
                    )

                    update_document(
                        doc["id"],
                        title=title.strip() or doc["title"],
                        category=category,
                        description=description.strip(),
                        tags=tags.strip(),
                        version=version.strip() or doc["version"],
                        uploaded_by=st.session_state.user_email,
                        pdf_bytes=pdf_bytes,
                    )

                    st.success("Document updated.")
                    st.rerun()

            with action_col2:
                pdf_bytes = load_pdf(doc["stored_filename"])

                if pdf_bytes:
                    st.download_button(
                        "Download",
                        data=pdf_bytes,
                        file_name=doc["filename"],
                        mime="application/pdf",
                        key=f"admin_download_{doc['id']}",
                        use_container_width=True,
                    )

            with action_col3:
                if st.button(
                    "Delete",
                    key=f"delete_{doc['id']}",
                    use_container_width=True,
                ):
                    st.session_state[f"confirm_delete_{doc['id']}"] = True

            if st.session_state.get(
                f"confirm_delete_{doc['id']}",
                False,
            ):
                st.warning(
                    "Deleting removes this document from the active "
                    "knowledge base."
                )

                if st.button(
                    "Confirm Delete",
                    key=f"confirm_delete_button_{doc['id']}",
                    type="primary",
                ):
                    delete_document(doc["id"])
                    st.session_state.pop(
                        f"confirm_delete_{doc['id']}",
                        None,
                    )
                    st.success("Document deleted.")
                    st.rerun()


# ============================================================
# ADMIN — ACCESS CODES
# ============================================================

def access_codes_page() -> None:
    if not st.session_state.is_admin:
        st.error("Administrator access required.")
        return

    top_header()

    st.markdown(
        '<div class="section-title">🔐 One-Time Access Codes</div>',
        unsafe_allow_html=True,
    )

    st.info(
        "Each generated code is stored as a SHA-256 hash and can be "
        "successfully consumed only once. Codes expire automatically."
    )

    if st.button(
        "＋ Generate New One-Time Access Code",
        type="primary",
    ):
        code = create_access_code(st.session_state.user_email)

        st.session_state.new_access_code = code

    if st.session_state.get("new_access_code"):
        st.success("New access code generated.")

        st.code(
            st.session_state.new_access_code,
            language=None,
        )

        st.warning(
            "Copy this code now. It is intentionally not stored in plain text "
            "and cannot be recovered after leaving this session."
        )

    connection = db()
    rows = connection.execute(
        """
        SELECT id, created_at, expires_at, used_at, revoked
        FROM access_codes
        ORDER BY id DESC
        LIMIT 50
        """
    ).fetchall()
    connection.close()

    if not rows:
        st.caption("No access codes have been generated.")
        return

    table = []

    for row in rows:
        if row["revoked"]:
            status = "Revoked"
        elif row["used_at"]:
            status = "Used"
        else:
            try:
                expiry = datetime.fromisoformat(row["expires_at"])
                if expiry.tzinfo is None:
                    expiry = expiry.replace(tzinfo=timezone.utc)
                status = "Active" if expiry > now_utc() else "Expired"
            except Exception:
                status = "Unknown"

        table.append(
            {
                "ID": row["id"],
                "Created": row["created_at"][:19].replace("T", " "),
                "Expires": row["expires_at"][:19].replace("T", " "),
                "Used": (
                    row["used_at"][:19].replace("T", " ")
                    if row["used_at"]
                    else "—"
                ),
                "Status": status,
            }
        )

    st.dataframe(
        pd.DataFrame(table),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Revoke Active Code")

    active_rows = [
        row for row in rows
        if not row["revoked"] and not row["used_at"]
    ]

    if active_rows:
        options = {
            f"Code #{row['id']} — expires {row['expires_at'][:19]}": row["id"]
            for row in active_rows
        }

        selected = st.selectbox(
            "Select code",
            list(options.keys()),
        )

        if st.button("Revoke Selected Code"):
            connection = db()
            connection.execute(
                "UPDATE access_codes SET revoked = 1 WHERE id = ?",
                (options[selected],),
            )
            connection.commit()
            connection.close()
            st.success("Code revoked.")
            st.rerun()
    else:
        st.caption("There are no active codes to revoke.")


# ============================================================
# ADMIN — ANALYTICS
# ============================================================

def analytics_page() -> None:
    if not st.session_state.is_admin:
        st.error("Administrator access required.")
        return

    top_header()

    docs = get_documents()

    st.markdown(
        '<div class="section-title">📊 Knowledge Base Analytics</div>',
        unsafe_allow_html=True,
    )

    if not docs:
        st.info("No documents to analyze yet.")
        return

    left, right = st.columns(2)

    with left:
        category_counts = (
            pd.DataFrame(docs)
            .groupby("category")
            .size()
            .reset_index(name="documents")
            .sort_values("documents", ascending=False)
        )

        st.markdown("#### Documents by Category")
        st.bar_chart(
            category_counts.set_index("category")
        )

    with right:
        views = pd.DataFrame(
            [
                {
                    "Document": doc["title"],
                    "Views": doc["view_count"],
                }
                for doc in docs
            ]
        ).sort_values("Views", ascending=False).head(10)

        st.markdown("#### Most Viewed Documents")
        st.bar_chart(
            views.set_index("Document")
        )

    st.markdown("#### Document Inventory")

    inventory = pd.DataFrame(
        [
            {
                "Title": d["title"],
                "Category": d["category"],
                "Version": d["version"],
                "Pages": d["page_count"],
                "Views": d["view_count"],
                "Updated": d["updated_at"][:10],
            }
            for d in docs
        ]
    )

    st.dataframe(
        inventory,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# MAIN ROUTER
# ============================================================

def main() -> None:
    if not st.session_state.authenticated:
        login_page()
        return

    sidebar()

    page = st.session_state.get("page", "Home")

    if page == "Home":
        home_page()
    elif page == "Browse All":
        browse_page()
    elif page == "Categories":
        categories_page()
    elif page == "Favorites":
        favorites_page()
    elif page == "Recent":
        recent_page()
    elif page == "Document":
        document_page()
    elif page == "Upload PDF":
        upload_page()
    elif page == "Manage Content":
        manage_content_page()
    elif page == "Access Codes":
        access_codes_page()
    elif page == "Analytics":
        analytics_page()
    else:
        home_page()


if __name__ == "__main__":
    main()
