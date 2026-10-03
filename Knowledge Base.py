import os
import re
import io
import base64
import hashlib
import sqlite3
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

try:
    from itsdangerous import URLSafeTimedSerializer
except Exception:
    URLSafeTimedSerializer = None

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ============================================================
# CONFIG
# ============================================================

APP_NAME = "Knowledge Base"
DATA_DIR = Path("knowledge_base_data")
PDF_DIR = DATA_DIR / "pdfs"
SOP_DIR = DATA_DIR / "sops"
DB_PATH = DATA_DIR / "knowledge_base.db"

DATA_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)
SOP_DIR.mkdir(exist_ok=True)

st.set_page_config(
    page_title="HPE Knowledge Base",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ============================================================
# CSS — visual language recreated from the uploaded reference
# ============================================================

st.markdown(
    """
<style>
:root{
 --navy:#06283a; --navy2:#0b4055; --teal:#00b894; --teal2:#00a982;
 --cyan:#00d4c7; --ink:#0b2940; --muted:#607786; --line:#d8e7e8;
 --glass:rgba(255,255,255,.92); --soft:#eef8f8; --blue:#1677df;
}
html,body,[class*="css"]{font-family:Arial,Helvetica,sans-serif}
.stApp{
 background:
 radial-gradient(circle at 4% 36%,rgba(0,196,173,.16),transparent 23%),
 radial-gradient(circle at 96% 22%,rgba(0,105,160,.14),transparent 26%),
 linear-gradient(180deg,#f7fbfc 0%,#edf7f7 100%);
 color:var(--ink);
}
[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"],
[data-testid="stStatusWidget"],[data-testid="stAppDeployButton"],
[data-testid="stMainMenu"],[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"]{display:none!important}
[data-testid="stHeader"]{height:0!important;min-height:0!important}
.block-container{max-width:1500px;padding:14px 30px 28px!important}

/* Hero */
.hero{
 position:relative; overflow:hidden; border-radius:0 0 26px 26px;
 min-height:255px; padding:22px 30px 28px; color:white;
 background:
 radial-gradient(ellipse at 75% 45%,rgba(0,210,198,.18),transparent 20%),
 linear-gradient(120deg,#03263a 0%,#06364d 48%,#052b3d 100%);
 box-shadow:0 10px 28px rgba(3,45,61,.15);
}
.hero:before{
 content:"";position:absolute;inset:0;opacity:.32;
 background:
 radial-gradient(ellipse at 62% 15%,transparent 0 22%,rgba(0,182,220,.45) 23%,transparent 24%),
 radial-gradient(ellipse at 80% 55%,transparent 0 17%,rgba(0,205,193,.38) 18%,transparent 19%);
 pointer-events:none;
}
.hero-brand{position:relative;z-index:2;width:190px}
.hero-mark{width:78px;height:11px;border:4px solid #00c89d;margin-bottom:7px}
.hero-brand-name{font-size:27px;font-weight:800;line-height:.92}
.hero-kb{font-size:19px;font-weight:700;margin-top:3px}
.hero-brand-copy{font-size:12px;line-height:1.35;margin-top:17px;color:#cfe4ea}
.hero-divider{position:absolute;left:234px;top:24px;height:76px;width:1px;background:rgba(255,255,255,.35)}
.hero-center{position:absolute;left:250px;right:245px;top:29px;text-align:center;z-index:2}
.hero-title{font-size:43px;font-weight:800;letter-spacing:-1.5px}
.hero-title span{color:#00d8bb}
.hero-sub{font-size:17px;color:#d8edf0;margin-top:4px}
.hero-right{position:absolute;right:32px;top:55px;width:155px;font-size:18px;font-weight:700;line-height:1.05;z-index:2}
.hero-right small{display:block;color:#a8d5d7;font-size:10px;font-weight:400;line-height:1.3;margin-top:10px}
.hero-search{position:absolute;left:24%;right:24%;top:116px;z-index:5}
.hero-search .stTextInput>div>div{
 border-radius:35px!important;background:#fff!important;border:0!important;
 box-shadow:0 7px 24px rgba(0,0,0,.22)!important;
}
.hero-search input{height:48px!important;font-size:15px!important;padding-left:20px!important}
.hero-search .stButton button{
 border-radius:50%!important;width:38px!important;height:38px!important;
 min-height:38px!important;padding:0!important;margin-top:5px!important;
 background:#00b894!important;border-color:#00b894!important;color:#fff!important;
}
.try-row{position:absolute;left:25%;right:25%;top:176px;z-index:4;text-align:center}
.try-label{font-size:10px;color:#d9eef1;margin-right:8px}
.try-pill{
 display:inline-block;padding:7px 13px;border-radius:18px;margin:2px 3px;
 background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.3);
 color:white;font-size:10px;
}

/* Category tiles */
.tile-grid{margin-top:12px}
.kb-tile{
 min-height:132px;background:rgba(255,255,255,.88);border:1px solid #d9e9e9;
 border-radius:16px;padding:16px;box-shadow:0 6px 20px rgba(11,58,75,.07);
}
.kb-icon{
 width:43px;height:43px;border-radius:11px;display:flex;align-items:center;
 justify-content:center;color:white;font-size:22px;font-weight:700;margin-bottom:10px;
}
.kb-tile-title{font-size:17px;font-weight:800;color:#0c2f47}
.kb-tile-desc{font-size:11px;color:#526b78;line-height:1.35;margin-top:4px}
.kb-arrow{float:right;width:28px;height:28px;border-radius:50%;background:#d9f7ef;
 color:#00a982;text-align:center;line-height:28px;font-size:18px}

/* Main cards */
.main-card{
 background:rgba(255,255,255,.9);border:1px solid #d7e8e8;border-radius:17px;
 box-shadow:0 7px 24px rgba(11,58,75,.06);padding:17px;
}
.section-head{font-size:18px;font-weight:800;color:#0b2d45}
.section-sub{font-size:11px;color:#6b7f8a;margin-top:2px}
.ai-card{
 background:linear-gradient(135deg,#f2ffff,#f5fbfb);
 border:1px solid #c9e8e5;border-radius:16px;padding:16px;
}
.ai-title{font-size:18px;font-weight:800;color:#0b2d45}
.ai-beta{font-size:9px;color:#058b75;background:#d9f7ef;border-radius:9px;padding:3px 6px;margin-left:5px}
.ai-answer{margin-top:10px;background:white;border:1px solid #dce9ec;border-radius:12px;padding:12px}
.ai-step{margin:7px 0;font-size:12px;line-height:1.45}
.step-no{
 display:inline-flex;width:20px;height:20px;border-radius:50%;align-items:center;
 justify-content:center;background:#12bd9a;color:white;font-weight:800;margin-right:8px;
}
.doc-card{
 background:#fff;border:1px solid #e0e9ec;border-radius:10px;padding:10px;height:100%;
}
.doc-thumb{
 height:82px;border-radius:7px;background:linear-gradient(145deg,#06283a,#0c5667);
 color:white;padding:10px;font-size:8px;font-weight:700;overflow:hidden;
}
.doc-title{font-size:11px;font-weight:800;line-height:1.2;margin-top:8px;color:#17354a}
.doc-meta{font-size:9px;color:#778b95;margin-top:5px}
.popular{
 background:#fff;border:1px solid #dfe9ec;border-radius:11px;padding:10px;
 min-height:63px;
}
.pop-title{font-size:11px;font-weight:800;color:#17354a}
.pop-meta{font-size:9px;color:#7a8d97;margin-top:2px}
.feature-link{font-size:11px;color:#056fc0;font-weight:700;text-align:right}

/* Results / reader */
.result-card{
 background:white;border:1px solid #dbe7ea;border-radius:10px;padding:10px;margin-bottom:7px;
}
.result-card.selected{border:1px solid #65d5bd;box-shadow:0 3px 12px rgba(0,169,130,.08)}
.result-label{font-size:9px;font-weight:800;color:#087c63}
.result-name{font-size:11px;font-weight:800;color:#12609a;line-height:1.25;margin-top:3px}
.result-snippet{font-size:9px;color:#5a707d;line-height:1.35;margin-top:5px}
.score{float:right;background:#e2f8f1;color:#087c63;border-radius:12px;padding:2px 6px;font-size:8px}
.reader{
 background:white;border:1px solid #dce7ea;border-radius:13px;padding:10px;
 box-shadow:0 5px 18px rgba(11,58,75,.05);
}
.reader-head{font-size:14px;font-weight:800;color:#102e45}
.reader-meta{font-size:9px;color:#72858e;margin-top:3px}
.reader-match{font-size:9px;color:#087c63;background:#e6f8f2;border-radius:5px;padding:5px 7px;display:inline-block;margin-top:5px}
.bottom-metric{
 background:rgba(255,255,255,.92);border:1px solid #dbe8ea;border-radius:15px;
 padding:11px 8px;text-align:center;box-shadow:0 4px 15px rgba(11,58,75,.05);
}
.metric-title{font-size:11px;font-weight:800;color:#17364b}
.metric-sub{font-size:8px;color:#71848d}

/* Admin */
.admin-shell{background:white;border:1px solid #d9e7ea;border-radius:16px;padding:20px;box-shadow:0 6px 20px rgba(11,58,75,.06)}
.admin-badge{display:inline-block;background:#e0f8f0;color:#087c63;border-radius:15px;padding:5px 10px;font-size:9px;font-weight:800}
.page-title{font-size:27px;font-weight:800;color:#0b2d45}
.page-sub{font-size:12px;color:#6c808a}
.sop-preview{background:#f7fbfb;border:1px solid #dce9eb;border-radius:10px;padding:12px;font-size:11px;color:#45606e;white-space:pre-wrap}

/* Streamlit controls */
button[kind="primary"]{background:#00a982!important;border-color:#00a982!important}
button[kind="primary"]:hover{background:#008e72!important}
.stButton button{border-radius:8px}
[data-testid="stFileUploader"]{background:white;border:1px dashed #a8c1c8;border-radius:10px}
[data-testid="stVerticalBlockBorderWrapper"]{border-color:#dce7ea!important;border-radius:12px!important}
.stDownloadButton button{border-color:#00a982!important;color:#087c63!important}
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
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS documents(
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
    CREATE TABLE IF NOT EXISTS chunks(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        page_number INTEGER NOT NULL,
        chunk_index INTEGER NOT NULL,
        text TEXT NOT NULL,
        FOREIGN KEY(document_id) REFERENCES documents(id)
    );
    CREATE INDEX IF NOT EXISTS idx_chunks_document_page ON chunks(document_id,page_number);
    CREATE INDEX IF NOT EXISTS idx_documents_category ON documents(category);
    """)
    conn.commit()
    conn.close()

init_db()

# ============================================================
# AUTH / ACCESS
# ============================================================

def get_admin_pin():
    try:
        pin = st.secrets.get("ADMIN_PIN")
        if pin:
            return str(pin)
    except Exception:
        pass
    return os.getenv("ADMIN_PIN","")

def get_access_code():
    try:
        code = st.secrets.get("ACCESS_CODE")
        if code:
            return str(code)
    except Exception:
        pass
    return os.getenv("ACCESS_CODE","")

ACCESS_CODE = get_access_code()
TOKEN_SECRET = str(os.getenv("TOKEN_SECRET","")).strip() or hashlib.sha256(
    f"{os.getcwd()}::{os.getenv('HOSTNAME','streamlit')}".encode()
).hexdigest()

def serializer():
    if URLSafeTimedSerializer is None:
        return None
    return URLSafeTimedSerializer(TOKEN_SECRET,salt="knowledge-base-browser-access")

def authorize():
    s=serializer()
    if not s:return False
    token=s.dumps({"authorized":True})
    st.query_params["kb_access"]=token
    st.session_state.access_authorized=True
    return True

def authorized():
    if st.session_state.get("access_authorized"): return True
    token=st.query_params.get("kb_access","")
    s=serializer()
    if s and token:
        try:
            st.session_state.access_authorized=bool(s.loads(token).get("authorized"))
        except Exception:
            pass
    return bool(st.session_state.get("access_authorized"))

if "access_authorized" not in st.session_state: st.session_state.access_authorized=False
if "page" not in st.session_state: st.session_state.page="Home"
if "search_query" not in st.session_state: st.session_state.search_query=""
if "selected_document" not in st.session_state: st.session_state.selected_document=None
if "viewer_page" not in st.session_state: st.session_state.viewer_page=1
if "admin_authenticated" not in st.session_state: st.session_state.admin_authenticated=False

def access_gate():
    st.markdown("""
    <div style="max-width:600px;margin:110px auto;text-align:center">
      <div class="hero-mark" style="margin:0 auto 8px"></div>
      <div style="font-size:18px;font-weight:800;color:#0b2d45">HPE Knowledge Base</div>
      <div style="font-size:28px;font-weight:800;color:#0b2d45;margin-top:18px">Secure Knowledge Access</div>
      <div style="font-size:13px;color:#6b7f8a;margin:8px 0 22px">Enter the configured access code to continue.</div>
    </div>
    """,unsafe_allow_html=True)
    if not ACCESS_CODE:
        st.error("Access control is not configured. Add ACCESS_CODE to Streamlit Secrets.")
        st.stop()
    with st.form("access"):
        code=st.text_input("Access Code",type="password")
        if st.form_submit_button("Access Knowledge Base",type="primary",use_container_width=True):
            if code.strip()==ACCESS_CODE:
                if authorize(): st.rerun()
            else: st.error("Invalid access code.")

if not authorized():
    access_gate()
    st.stop()

# ============================================================
# DOCUMENT HELPERS
# ============================================================

def clean_text(text):
    text=text.replace("\x00"," ")
    text=re.sub(r"[ \t]+"," ",text)
    text=re.sub(r"\n{3,}","\n\n",text)
    return text.strip()

def make_hash(data): return hashlib.sha256(data).hexdigest()

def split_text(text,chunk_size=1100,overlap=180):
    words=text.split()
    if not words:return []
    out=[]; start=0
    while start<len(words):
        end=min(len(words),start+chunk_size)
        part=" ".join(words[start:end]).strip()
        if part:out.append(part)
        if end>=len(words):break
        start=max(0,end-overlap)
    return out

def extract_pdf(data):
    pages=[]
    with fitz.open(stream=data,filetype="pdf") as doc:
        for n,p in enumerate(doc,start=1):
            pages.append((n,clean_text(p.get_text("text"))))
        return pages,len(doc)

def document_exists(file_hash):
    conn=db()
    row=conn.execute("SELECT id FROM documents WHERE file_hash=?",(file_hash,)).fetchone()
    conn.close()
    return row

def add_document(filename,data,category="General"):
    h=make_hash(data)
    if document_exists(h): return False,"This PDF has already been uploaded."
    try: pages,page_count=extract_pdf(data)
    except Exception as e: return False,f"Could not read PDF: {e}"
    safe=re.sub(r"[^A-Za-z0-9._-]+","_",filename)
    path=PDF_DIR/f"{h[:12]}_{safe}"
    path.write_bytes(data)
    conn=db(); now=datetime.now().isoformat(timespec="seconds")
    cur=conn.execute("""INSERT INTO documents
      (filename,stored_path,file_hash,category,page_count,file_size,uploaded_at,indexed_at,status)
      VALUES(?,?,?,?,?,?,?,?,?)""",
      (filename,str(path),h,category,page_count,len(data),now,now,"Indexed"))
    doc_id=cur.lastrowid
    for page,text in pages:
        for idx,chunk in enumerate(split_text(text)):
            conn.execute("INSERT INTO chunks(document_id,page_number,chunk_index,text) VALUES(?,?,?,?)",
                         (doc_id,page,idx,chunk))
    conn.commit(); conn.close()
    return True,f"{filename} indexed successfully."

def get_documents():
    conn=db()
    rows=[dict(r) for r in conn.execute("SELECT * FROM documents ORDER BY uploaded_at DESC").fetchall()]
    conn.close(); return rows

def get_categories():
    conn=db()
    rows=[r["category"] for r in conn.execute("SELECT DISTINCT category FROM documents ORDER BY category").fetchall()]
    conn.close(); return rows

def get_all_chunks():
    conn=db()
    rows=[dict(r) for r in conn.execute("""SELECT c.*,d.filename,d.category,d.stored_path
        FROM chunks c JOIN documents d ON d.id=c.document_id ORDER BY c.id""").fetchall()]
    conn.close(); return rows

def get_document(doc_id):
    conn=db(); row=conn.execute("SELECT * FROM documents WHERE id=?",(doc_id,)).fetchone(); conn.close()
    return dict(row) if row else None

def delete_document(doc_id):
    conn=db()
    row=conn.execute("SELECT stored_path FROM documents WHERE id=?",(doc_id,)).fetchone()
    if row:
        try: Path(row["stored_path"]).unlink(missing_ok=True)
        except Exception: pass
    conn.execute("DELETE FROM chunks WHERE document_id=?",(doc_id,))
    conn.execute("DELETE FROM documents WHERE id=?",(doc_id,))
    conn.commit(); conn.close()

def format_bytes(n):
    if not n:return "0 B"
    n=float(n)
    if n<1024:return f"{n:.0f} B"
    if n<1024**2:return f"{n/1024:.1f} KB"
    if n<1024**3:return f"{n/1024**2:.1f} MB"
    return f"{n/1024**3:.1f} GB"

# ============================================================
# SEARCH
# ============================================================

@st.cache_data(ttl=60,show_spinner=False)
def build_index():
    rows=get_all_chunks()
    if not rows:return None,None,[]
    texts=[r["text"] for r in rows]
    vec=TfidfVectorizer(lowercase=True,stop_words="english",ngram_range=(1,2),
                        min_df=1,max_df=.98,sublinear_tf=True,dtype="float32")
    return vec,vec.fit_transform(texts),rows

def snippet(text,query,radius=250):
    t=re.sub(r"\s+"," ",text).strip()
    if not query:return t[:radius]+("..." if len(t)>radius else "")
    terms=[x.lower() for x in re.findall(r"\w+",query) if len(x)>2]
    low=t.lower(); pos=[low.find(x) for x in terms if low.find(x)>=0]
    if not pos:return t[:radius]+("..." if len(t)>radius else "")
    center=min(pos); start=max(0,center-radius//2); end=min(len(t),start+radius)
    return ("..." if start else "")+t[start:end]+("..." if end<len(t) else "")

@st.cache_data(ttl=30,show_spinner=False)
def search_documents(query,category="All Categories",top_k=10):
    vec,matrix,rows=build_index()
    if vec is None:return []
    scores=cosine_similarity(vec.transform([query.strip()]),matrix).flatten()
    best={}
    for i,score in enumerate(scores):
        r=rows[i]
        if category!="All Categories" and r["category"]!=category:continue
        if score<=0:continue
        item={**r,"score":float(score),"snippet":snippet(r["text"],query)}
        old=best.get(r["document_id"])
        if old is None or item["score"]>old["score"]:best[r["document_id"]]=item
    return sorted(best.values(),key=lambda x:x["score"],reverse=True)[:top_k]

@st.cache_data(ttl=300,show_spinner=False)
def render_page(path,page,scale=1.5,query=""):
    try:
        with fitz.open(path) as pdf:
            p=pdf[page-1]
            terms=[x for x in re.findall(r"[A-Za-z0-9]+",query) if len(x)>2]
            for term in terms[:12]:
                try:
                    for rect in p.search_for(term):
                        a=p.add_highlight_annot(rect); a.update()
                except Exception: pass
            pix=p.get_pixmap(matrix=fitz.Matrix(scale,scale),alpha=False)
            return pix.tobytes("png")
    except Exception:return None

def find_matches(path,query,limit=8):
    terms=[x.lower() for x in re.findall(r"[A-Za-z0-9]+",query) if len(x)>2]
    out=[]
    try:
        with fitz.open(path) as pdf:
            for n,p in enumerate(pdf,start=1):
                raw=clean_text(p.get_text("text")); low=raw.lower()
                score=sum(low.count(t) for t in terms)
                if score:out.append({"page":n,"score":score,"snippet":snippet(raw,query,190)})
    except Exception:return []
    return sorted(out,key=lambda x:(-x["score"],x["page"]))[:limit]

def clear_caches():
    for fn in (build_index,search_documents,render_page):
        try:fn.clear()
        except Exception:pass

# ============================================================
# SOP CREATOR — admin can create a PDF SOP and immediately index it
# ============================================================

def create_sop_pdf(title,category,owner,content):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is not installed. Add reportlab to requirements.txt.")
    safe=re.sub(r"[^A-Za-z0-9._-]+","_",title).strip("_") or "HPE_SOP"
    out=SOP_DIR/f"{safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    styles=getSampleStyleSheet()
    title_style=ParagraphStyle("SOPTitle",parent=styles["Title"],fontSize=22,leading=27,textColor=colors.HexColor("#0b2d45"),alignment=TA_CENTER,spaceAfter=12)
    h_style=ParagraphStyle("SOPHead",parent=styles["Heading2"],fontSize=14,leading=18,textColor=colors.HexColor("#087c63"),spaceBefore=10,spaceAfter=6)
    body_style=ParagraphStyle("SOPBody",parent=styles["BodyText"],fontSize=9.5,leading=14,textColor=colors.HexColor("#263f50"),spaceAfter=7)
    meta_style=ParagraphStyle("SOPMeta",parent=styles["BodyText"],fontSize=8,color=colors.HexColor("#657983"),leading=11)
    doc=SimpleDocTemplate(str(out),pagesize=letter,rightMargin=45,leftMargin=45,topMargin=45,bottomMargin=45)
    story=[
      Paragraph(title,title_style),
      Paragraph(f"<b>Category:</b> {category} &nbsp;&nbsp; <b>Owner:</b> {owner or 'Knowledge Base Admin'}",meta_style),
      Paragraph(f"<b>Created:</b> {datetime.now().strftime('%B %d, %Y')}",meta_style),
      Spacer(1,14)
    ]
    for block in re.split(r"\n\s*\n",content.strip()):
        block=block.strip()
        if not block:continue
        first=block.splitlines()[0].strip()
        if re.match(r"^(?:\d+[\.\)]\s+|[A-Z][A-Za-z0-9 &/&-]{2,60}:$)",first):
            story.append(Paragraph(first.rstrip(":"),h_style))
            rest="\n".join(block.splitlines()[1:]).strip()
            if rest: story.append(Paragraph(rest.replace("\n","<br/>"),body_style))
        else:
            story.append(Paragraph(block.replace("\n","<br/>"),body_style))
    doc.build(story)
    return out

# ============================================================
# HEADER / HERO
# ============================================================

def go(page):
    st.session_state.page=page
    st.rerun()

st.markdown("""
<div class="hero">
  <div class="hero-brand">
    <div class="hero-mark"></div>
    <div class="hero-brand-name">HPE</div>
    <div class="hero-kb">Knowledge Base</div>
    <div class="hero-brand-copy">Find answers. Solve faster.<br>Power your support with HPE knowledge.</div>
  </div>
  <div class="hero-divider"></div>
  <div class="hero-center">
    <div class="hero-title">How can we help you <span>today?</span></div>
    <div class="hero-sub">Search HPE documentation, guides, and solutions with AI.</div>
  </div>
  <div class="hero-right">Accelerating<br>what's next<br>together<small>Built for fast support discovery</small></div>
  <div class="try-row">
    <span class="try-label">Try asking:</span>
    <span class="try-pill">How to renew a license?</span>
    <span class="try-pill">ClearPass troubleshooting</span>
    <span class="try-pill">iLO configuration</span>
    <span class="try-pill">Aruba switch setup</span>
    <span class="try-pill">Gen11 firmware update</span>
  </div>
</div>
""",unsafe_allow_html=True)

# Hero search is a real Streamlit input, not part of the image.
with st.container():
    hq,hb=st.columns([6.4,.55],gap="small")
    with hq:
        hero_query=st.text_input("Hero search",value=st.session_state.search_query,
            placeholder="Ask a question or search for a document...",label_visibility="collapsed",key="hero_query")
    with hb:
        hero_submit=st.button("➜",key="hero_submit",type="primary",use_container_width=True)
if hero_submit:
    st.session_state.search_query=hero_query.strip()
    st.session_state.page="Search"
    st.rerun()

# ============================================================
# CATEGORY TILES
# ============================================================

categories=[
 ("🖥","Compute","Servers, iLO, Gen10/Gen11, firmware, hardware guides","#00b894"),
 ("⌘","Networking","Aruba, Switches, Controllers, Access Points","#1677df"),
 ("◉","Storage","Alletra, Nimble, 3PAR, Primera, StoreOnce","#7c4de8"),
 ("▤","Software & Licensing","Licensing, Subscriptions, Entitlements","#f28b32"),
 ("♢","Security","ClearPass, NAC, Policy Management","#ed4b58"),
 ("⚙","Support & Tools","HPE Tools, Portals, Utilities, Best Practices","#607b8c"),
]
cols=st.columns(6,gap="small")
for i,(icon,title,desc,bg) in enumerate(categories):
    with cols[i]:
        st.markdown(f"""<div class="kb-tile">
          <div class="kb-icon" style="background:{bg}">{icon}</div>
          <div class="kb-arrow">›</div>
          <div class="kb-tile-title">{title}</div>
          <div class="kb-tile-desc">{desc}</div>
        </div>""",unsafe_allow_html=True)
        if st.button("Explore",key=f"cat_{i}",use_container_width=True):
            st.session_state.search_query=title
            st.session_state.page="Search"
            st.rerun()

# ============================================================
# HOME: AI ASSISTANT + FEATURED DOCS + POPULAR TOPICS
# ============================================================

docs=get_documents()
if st.session_state.page=="Home":
    left,right=st.columns([1.02,1.0],gap="medium")
    with left:
        st.markdown("""<div class="ai-card">
          <div class="ai-title">🤖 HPE AI Assistant <span class="ai-beta">BETA</span></div>
          <div style="font-size:11px;color:#667c87;margin-top:3px">Get instant answers grounded in your indexed HPE documentation.</div>
        </div>""",unsafe_allow_html=True)
        q=st.text_input("AI question",placeholder="How do I troubleshoot ClearPass licensing issues?",label_visibility="collapsed",key="home_ai_q")
        if st.button("Ask AI",type="primary",use_container_width=True,key="ask_ai_home"):
            st.session_state.search_query=q.strip()
            st.session_state.page="Search"
            st.rerun()
        if docs:
            sample=search_documents("ClearPass licensing",top_k=1)
            if sample:
                s=sample[0]
                st.markdown(f"""<div class="ai-answer">
                  <div style="font-size:10px;color:#087c63;font-weight:800">GROUNDED ANSWER · {s['filename']}</div>
                  <div style="font-size:12px;line-height:1.5;margin-top:7px">{s['snippet']}</div>
                  <div style="margin-top:8px;font-size:9px;color:#6f828c">This assistant retrieves and cites indexed document content; it does not invent unsupported procedures.</div>
                </div>""",unsafe_allow_html=True)
            else:
                st.info("Upload or create SOPs to activate document-grounded answers.")
        else:
            st.info("Your AI Assistant will become useful as soon as PDFs are indexed.")
    with right:
        st.markdown("""<div class="main-card">
          <div class="section-head">⭐ Featured Documents</div>
          <div class="section-sub">Most useful and commonly accessed resources.</div>
        </div>""",unsafe_allow_html=True)
        fdocs=docs[:4]
        if fdocs:
            fcols=st.columns(min(4,len(fdocs)),gap="small")
            for i,d in enumerate(fdocs):
                with fcols[i]:
                    st.markdown(f"""<div class="doc-card">
                      <div class="doc-thumb">HPE<br><br>{d['filename'][:42]}</div>
                      <div class="doc-title">{d['filename']}</div>
                      <div class="doc-meta">PDF · {format_bytes(d['file_size'])}</div>
                    </div>""",unsafe_allow_html=True)
                    if st.button("View",key=f"view_feature_{d['id']}",use_container_width=True):
                        st.session_state.selected_document=d["id"]; st.session_state.viewer_page=1; st.session_state.page="Reader"; st.rerun()
        else:
            st.caption("No documents yet. Use Manage Documents from the gear menu.")
    st.write("")
    pc1,pc2=st.columns([1,1],gap="medium")
    with pc1:
        st.markdown("""<div class="main-card"><div class="section-head">🔥 Popular Topics</div>
        <div class="section-sub">Search shortcuts for common HPE support themes.</div></div>""",unsafe_allow_html=True)
        topics=[
          ("🔑","License Activation","Search licensing and entitlement documents"),
          ("🔐","ClearPass Troubleshooting","Search NAC, authentication and policy content"),
          ("🖥","iLO Configuration","Search remote management and server setup"),
          ("♻","Firmware Updates","Search firmware and update guidance"),
          ("⌘","Switch Setup","Search Aruba switching and configuration"),
          ("◉","Alletra Storage","Search HPE storage architecture and operations"),
        ]
        for r in range(3):
            a,b=st.columns(2,gap="small")
            for col,j in zip((a,b),(r*2,r*2+1)):
                icon,title,desc=topics[j]
                with col:
                    st.markdown(f"""<div class="popular"><span style="font-size:17px">{icon}</span>
                    <span class="pop-title">{title}</span><div class="pop-meta">{desc}</div></div>""",unsafe_allow_html=True)
                    if st.button("Search",key=f"topic_{j}",use_container_width=True):
                        st.session_state.search_query=title; st.session_state.page="Search"; st.rerun()
    with pc2:
        st.markdown("""<div class="main-card"><div class="section-head">🚀 Knowledge Base Status</div>
        <div class="section-sub">Your local indexed knowledge layer.</div></div>""",unsafe_allow_html=True)
        total_pages=sum(int(d["page_count"] or 0) for d in docs)
        m=st.columns(3)
        for col,val,label in zip(m,(len(docs),total_pages,len(get_categories())),("Documents","Indexed pages","Categories")):
            with col:
                st.markdown(f"""<div class="bottom-metric"><div style="font-size:22px;font-weight:800;color:#00a982">{val:,}</div>
                <div class="metric-title">{label}</div><div class="metric-sub">Ready for search</div></div>""",unsafe_allow_html=True)

# ============================================================
# SEARCH
# ============================================================

elif st.session_state.page=="Search":
    st.markdown('<div class="main-card">',unsafe_allow_html=True)
    q1,q2,q3=st.columns([5.8,1.2,1.1],gap="small")
    with q1:
        query=st.text_input("Search knowledge",value=st.session_state.search_query,
                            placeholder="Ask a question or search for a document...",label_visibility="collapsed",key="search_box")
    with q2:
        cat=st.selectbox("Category",["All Categories"]+get_categories(),label_visibility="collapsed",key="search_cat")
    with q3:
        submit=st.button("Search",type="primary",use_container_width=True,key="search_btn")
    st.markdown('</div>',unsafe_allow_html=True)
    if submit:
        st.session_state.search_query=query.strip()
    active=st.session_state.search_query.strip()
    if active:
        results=search_documents(active,cat,10)
        if not results:
            st.warning("No matching document was found. Try different keywords or upload another PDF.")
        else:
            l,r=st.columns([.82,1.8],gap="medium")
            with l:
                st.markdown('<div class="section-head">Search Results</div>',unsafe_allow_html=True)
                for i,res in enumerate(results):
                    selected=st.session_state.selected_document==res["document_id"]
                    st.markdown(f"""<div class="result-card {'selected' if selected else ''}">
                      <span class="result-label">{'HIGHEST MATCH' if i==0 else 'SEARCH RESULT'}</span>
                      <span class="score">{min(99,max(1,round(res['score']*100)))}%</span>
                      <div class="result-name">{res['filename']}</div>
                      <div style="font-size:8px;color:#7a8b93">Page {res['page_number']} · {res['category']}</div>
                      <div class="result-snippet">{res['snippet']}</div>
                    </div>""",unsafe_allow_html=True)
                    if st.button("Open Result",key=f"open_{res['document_id']}",use_container_width=True):
                        st.session_state.selected_document=res["document_id"]
                        st.session_state.viewer_page=int(res["page_number"])
                        st.session_state.page="Reader"
                        st.rerun()
            with r:
                best=results[0]
                doc=get_document(best["document_id"])
                page=max(1,min(int(st.session_state.viewer_page or best["page_number"]),int(doc["page_count"] or 1)))
                st.session_state.viewer_page=page
                st.markdown(f"""<div class="reader"><div class="reader-head">▣ {doc['filename']}</div>
                <div class="reader-meta">Page {page} of {doc['page_count']} · Search: “{active}”</div>
                <div class="reader-match">Selected source · matching text is highlighted in the PDF</div></div>""",unsafe_allow_html=True)
                nav1,nav2,nav3=st.columns([1,4,1],gap="small")
                with nav1:
                    if st.button("‹",disabled=page<=1,key="prev_search",use_container_width=True):
                        st.session_state.viewer_page=page-1; st.rerun()
                with nav2:
                    zoom=st.select_slider("Zoom",[100,125,150,175],value=125,label_visibility="collapsed",key="search_zoom")
                with nav3:
                    if st.button("›",disabled=page>=doc["page_count"],key="next_search",use_container_width=True):
                        st.session_state.viewer_page=page+1; st.rerun()
                image=render_page(doc["stored_path"],page,scale=float(zoom)/100*1.05,query=active)
                if image: st.image(image,use_container_width=True)
                matches=find_matches(doc["stored_path"],active)
                if matches:
                    st.markdown("**Matches in this PDF**")
                    for m in matches:
                        if st.button(f"Page {m['page']} · {m['score']} match(es)",key=f"match_{doc['id']}_{m['page']}",use_container_width=True):
                            st.session_state.viewer_page=m["page"]; st.rerun()
                        st.caption(m["snippet"])

# ============================================================
# READER
# ============================================================

elif st.session_state.page=="Reader":
    doc=get_document(st.session_state.selected_document) if st.session_state.selected_document else None
    if not doc:
        st.warning("Document not found."); st.session_state.page="Home"; st.rerun()
    st.markdown(f'<div class="page-title">Document Reader</div><div class="page-sub">{doc["filename"]} · {doc["category"]}</div>',unsafe_allow_html=True)
    page=max(1,min(int(st.session_state.viewer_page),int(doc["page_count"] or 1)))
    a,b,c=st.columns([1,5,1])
    with a:
        if st.button("← Previous",disabled=page<=1,use_container_width=True):
            st.session_state.viewer_page=page-1; st.rerun()
    with b:
        st.markdown(f"<div style='text-align:center;font-size:11px;color:#607786;padding:8px'>Page <b>{page}</b> / {doc['page_count']}</div>",unsafe_allow_html=True)
    with c:
        if st.button("Next →",disabled=page>=doc["page_count"],use_container_width=True):
            st.session_state.viewer_page=page+1; st.rerun()
    image=render_page(doc["stored_path"],page,scale=1.35)
    if image: st.image(image,use_container_width=True)
    if st.button("← Back to Knowledge Base",use_container_width=True): go("Home")

# ============================================================
# ADMIN LOGIN / MANAGEMENT
# ============================================================

elif st.session_state.page=="Admin Login":
    st.markdown('<div class="page-title">Admin Access</div><div class="page-sub">Manage the document library and create new SOP PDFs.</div>',unsafe_allow_html=True)
    pin=st.text_input("Admin PIN",type="password",key="admin_pin")
    if st.button("Unlock",type="primary",use_container_width=True):
        if get_admin_pin() and pin==get_admin_pin():
            st.session_state.admin_authenticated=True; st.session_state.page="Manage Documents"; st.rerun()
        else: st.error("Incorrect admin PIN.")
    if st.button("← Back to Knowledge Base",use_container_width=True): go("Home")

elif st.session_state.page=="Manage Documents":
    if not st.session_state.admin_authenticated: go("Admin Login")
    st.markdown('<span class="admin-badge">ADMIN ONLY</span>',unsafe_allow_html=True)
    st.markdown('<div class="page-title">Manage Documents</div><div class="page-sub">Upload PDFs, create SOPs, and keep the knowledge base indexed.</div>',unsafe_allow_html=True)

    upload_col,sop_col=st.columns([1,1],gap="medium")
    with upload_col:
        st.markdown('<div class="admin-shell"><div class="section-head">📄 Upload PDF</div><div class="section-sub">Any readable PDF becomes searchable immediately.</div></div>',unsafe_allow_html=True)
        cat=st.selectbox("Category",["General","Policies","Procedures","Technical Support","Licensing","Training","Product","Account Management","Security","Networking","Storage","Compute","Other"],key="admin_cat")
        files=st.file_uploader("Upload PDF files",type=["pdf"],accept_multiple_files=True,key="pdfs")
        if st.button("Upload & Index",type="primary",use_container_width=True,key="upload_index"):
            if not files: st.warning("Select at least one PDF.")
            else:
                for f in files:
                    ok,msg=add_document(f.name,f.getvalue(),cat)
                    (st.success if ok else st.warning)(msg)
                clear_caches(); st.rerun()
    with sop_col:
        st.markdown('<div class="admin-shell"><div class="section-head">📝 Create SOP</div><div class="section-sub">Create a formatted PDF SOP directly inside the admin page, then automatically add it to the knowledge base.</div></div>',unsafe_allow_html=True)
        title=st.text_input("SOP title",placeholder="Example: HPE ProLiant Gen11 Basic Support Procedure",key="sop_title")
        sop_cat=st.selectbox("SOP category",["Procedures","Technical Support","Compute","Networking","Storage","Security","Licensing","Training","Other"],key="sop_cat")
        owner=st.text_input("Owner / team",placeholder="Knowledge Base Team",key="sop_owner")
        content=st.text_area("SOP content",height=260,placeholder="Purpose\\n\\nDescribe the procedure...\\n\\nPrerequisites\\n\\n1. ...\\n2. ...\\n\\nProcedure\\n\\n1. ...",key="sop_content")
        if st.button("Create SOP PDF & Add to Knowledge Base",type="primary",use_container_width=True,key="create_sop"):
            if not title.strip() or not content.strip(): st.warning("Enter an SOP title and content.")
            else:
                try:
                    path=create_sop_pdf(title,sop_cat,owner,content)
                    ok,msg=add_document(path.name,path.read_bytes(),sop_cat)
                    if ok:
                        st.success(msg)
                        st.download_button("Download created SOP",data=path.read_bytes(),file_name=path.name,mime="application/pdf",use_container_width=True)
                        clear_caches()
                    else: st.warning(msg)
                except Exception as e: st.error(str(e))

    st.markdown("---")
    st.markdown('<div class="section-head">Document Library</div>',unsafe_allow_html=True)
    docs=get_documents()
    if not docs: st.info("No documents indexed yet.")
    for d in docs:
        with st.container(border=True):
            a,b,c=st.columns([5,2,1])
            with a: st.markdown(f"**📄 {d['filename']}**"); st.caption(f"{d['category']} · {d['page_count']} pages · {format_bytes(d['file_size'])} · Indexed {d['indexed_at']}")
            with b:
                if st.button("Open",key=f"admin_open_{d['id']}",use_container_width=True):
                    st.session_state.selected_document=d["id"]; st.session_state.viewer_page=1; st.session_state.page="Reader"; st.rerun()
            with c:
                if st.button("Delete",key=f"admin_delete_{d['id']}",use_container_width=True):
                    delete_document(d["id"]); clear_caches(); st.rerun()
    if st.button("← Back to Knowledge Base",use_container_width=True): go("Home")

# ============================================================
# TOP / BOTTOM NAVIGATION
# ============================================================

st.markdown("<div style='height:18px'></div>",unsafe_allow_html=True)
n1,n2,n3,n4=st.columns(4,gap="small")
with n1:
    if st.button("⌂ Home",use_container_width=True): go("Home")
with n2:
    if st.button("⌕ Search",use_container_width=True): go("Search")
with n3:
    if st.button("📄 Documents",use_container_width=True): go("Reader" if st.session_state.selected_document else "Home")
with n4:
    if st.button("⚙ Admin",use_container_width=True): go("Manage Documents" if st.session_state.admin_authenticated else "Admin Login")
st.markdown("<div style='text-align:center;color:#71858e;font-size:9px;padding:7px'>HPE Knowledge Base · Document-grounded search · Admin-managed PDF library</div>",unsafe_allow_html=True)
