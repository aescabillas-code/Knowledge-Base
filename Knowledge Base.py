import os
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from html import escape

import streamlit as st

try:
    import fitz
except Exception:
    fitz = None

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
except Exception:
    TfidfVectorizer = None
    cosine_similarity = None


# ============================================================
# PAGE
# ============================================================
st.set_page_config(
    page_title="HPE Knowledge Base",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path("knowledge_base_data")
PDF_DIR = BASE_DIR / "pdfs"
DB_PATH = BASE_DIR / "knowledge_base.db"
RESET_MARKER = BASE_DIR / ".knowledge_base_v2_reset_complete"

BASE_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)


# ============================================================
# SIX PRIMARY PRODUCT GROUPS
# The home page intentionally has EXACTLY the six tile groups
# visible in the supplied reference image. Each group opens a
# complete product/topic/source directory.
# ============================================================
PRODUCT_GROUPS = {
    "Compute": {
        "icon": "▣",
        "class": "mint",
        "description": "Servers, iLO, Gen10/Gen11, firmware, hardware guides",
        "products": [
            "HPE ProLiant Compute",
            "HPE ProLiant Gen10",
            "HPE ProLiant Gen11",
            "HPE ProLiant Gen12",
            "HPE iLO",
            "HPE Compute Ops Management",
            "HPE OneView",
            "HPE Superdome Flex",
        ],
        "topics": [
            "Server Setup & Configuration",
            "iLO Configuration",
            "BIOS / UEFI",
            "Firmware Updates",
            "Hardware Health",
            "POST & Boot Troubleshooting",
            "Storage Controllers",
            "NIC & Network Adapters",
            "Power & Thermal",
            "Drivers",
            "Lifecycle Management",
            "Server Monitoring",
        ],
        "sources": [
            ("HPE Compute", "https://www.hpe.com/us/en/compute.html"),
            ("HPE ProLiant", "https://www.hpe.com/us/en/compute/proliant.html"),
            ("HPE Support Center", "https://support.hpe.com/"),
            ("HPE Documentation", "https://www.hpe.com/info/docs"),
        ],
    },
    "Networking": {
        "icon": "⌘",
        "class": "blue",
        "description": "Aruba, switches, controllers, access points, routing",
        "products": [
            "HPE Networking",
            "HPE Aruba Networking",
            "Aruba CX Switches",
            "Aruba Access Points",
            "Aruba Gateways",
            "Aruba Central",
            "Aruba AirWave",
            "Juniper Networking",
        ],
        "topics": [
            "Switch Setup",
            "VLAN Configuration",
            "Trunking",
            "Routing",
            "STP",
            "Port Troubleshooting",
            "Wireless Client Connectivity",
            "SSID Configuration",
            "Access Point Troubleshooting",
            "Aruba Central",
            "Gateway & SD-Branch",
            "Network Monitoring",
        ],
        "sources": [
            ("HPE Networking", "https://www.hpe.com/us/en/products/networking.html"),
            ("HPE Aruba Networking", "https://www.hpe.com/us/en/hpe-aruba-networking.html"),
            ("Aruba Networking Documentation", "https://arubanetworking.hpe.com/techdocs/"),
            ("HPE Support Center", "https://support.hpe.com/"),
        ],
    },
    "Storage": {
        "icon": "▤",
        "class": "violet",
        "description": "Alletra, MSA, Primera, StoreOnce, data protection",
        "products": [
            "HPE Alletra Storage",
            "HPE Alletra Storage MP B10000",
            "HPE Alletra Storage MP X10000",
            "HPE Alletra 5000",
            "HPE Alletra 6000",
            "HPE Alletra 9000",
            "HPE MSA",
            "HPE Primera",
            "HPE StoreOnce",
        ],
        "topics": [
            "Storage Setup",
            "Host Connectivity",
            "Volume Mapping",
            "Multipathing",
            "Capacity",
            "Performance",
            "Snapshots",
            "Replication",
            "Data Protection",
            "Storage Networking",
            "Alerts & Health",
            "Backup & Recovery",
        ],
        "sources": [
            ("HPE Alletra Storage", "https://www.hpe.com/us/en/products/storage/alletra.html"),
            ("HPE Storage", "https://www.hpe.com/us/en/storage.html"),
            ("HPE Support Center", "https://support.hpe.com/"),
            ("HPE Documentation", "https://www.hpe.com/info/docs"),
        ],
    },
    "Software & Licensing": {
        "icon": "▥",
        "class": "orange",
        "description": "Licensing, subscriptions, entitlements and software",
        "products": [
            "HPE Software",
            "HPE GreenLake Services",
            "HPE iLO Licensing",
            "HPE OneView",
            "HPE Compute Ops Management",
            "HPE Aruba Networking Subscriptions",
            "HPE Storage Software",
            "HPE Data Services",
        ],
        "topics": [
            "License Activation",
            "License Renewal",
            "Entitlements",
            "Subscriptions",
            "License Assignment",
            "License Consumption",
            "Activation & Registration",
            "Portal Access",
            "Software Downloads",
            "Subscription Management",
            "Licensing Errors",
            "Support Entitlements",
        ],
        "sources": [
            ("HPE Software", "https://buy.hpe.com/us/en/software"),
            ("HPE GreenLake", "https://www.hpe.com/us/en/greenlake.html"),
            ("HPE Support Center", "https://support.hpe.com/"),
            ("HPE Store", "https://buy.hpe.com/"),
        ],
    },
    "Security": {
        "icon": "♢",
        "class": "red",
        "description": "ClearPass, NAC, policy management, zero trust",
        "products": [
            "HPE Aruba Networking ClearPass",
            "ClearPass Policy Manager",
            "ClearPass Guest",
            "ClearPass OnGuard",
            "Aruba Policy Manager",
            "Network Access Control",
            "HPE Aruba Networking SSE",
            "HPE Security",
        ],
        "topics": [
            "ClearPass Troubleshooting",
            "RADIUS",
            "802.1X",
            "MAC Authentication",
            "Role Mapping",
            "Enforcement Policies",
            "Access Tracker",
            "Guest Access",
            "OnGuard",
            "Certificates",
            "Zero Trust",
            "SASE",
        ],
        "sources": [
            ("ClearPass Documentation", "https://arubanetworking.hpe.com/techdocs/ClearPass/"),
            ("HPE Aruba Networking", "https://www.hpe.com/us/en/hpe-aruba-networking.html"),
            ("HPE Security", "https://www.hpe.com/us/en/security.html"),
            ("HPE Support Center", "https://support.hpe.com/"),
        ],
    },
    "Support & Tools": {
        "icon": "⚙",
        "class": "slate",
        "description": "HPE tools, portals, utilities, support and best practices",
        "products": [
            "HPE Support Center",
            "HPE InfoSight",
            "HPE GreenLake",
            "HPE Service Portal",
            "HPE OneView",
            "HPE iLO",
            "HPE Active Health System",
            "HPE Smart Update Manager",
            "HPE Documentation",
        ],
        "topics": [
            "Case Troubleshooting",
            "Support Case Management",
            "Technical Escalation",
            "Documentation Search",
            "Knowledge Articles",
            "Logs & Diagnostics",
            "Health Monitoring",
            "Support Entitlements",
            "Warranty & Care Packs",
            "Best Practices",
            "Customer Communication",
            "Problem Management",
        ],
        "sources": [
            ("HPE Support Center", "https://support.hpe.com/"),
            ("HPE Documentation", "https://www.hpe.com/info/docs"),
            ("HPE Services", "https://www.hpe.com/us/en/services.html"),
            ("HPE GreenLake", "https://www.hpe.com/us/en/greenlake.html"),
        ],
    },
}


# ============================================================
# AI-READY SEED RECORDS
# One question = one self-contained answer.
# ============================================================
SEED_KB = [
    ("CPPM-001","Security","ClearPass Policy Manager",
     "How do I troubleshoot a ClearPass authentication failure?",
     "Check the request, selected ClearPass service, authentication source, identity result, role mapping, and enforcement result. Start with the exact failure reason before changing configuration.",
     "1. Identify the affected request and timestamp.\n2. Check which service processed the request.\n3. Confirm the authentication source is reachable.\n4. Review the exact authentication failure reason.\n5. If authentication succeeds, inspect role mapping and enforcement.\n6. Retest with a controlled request.",
     "ClearPass, authentication failure, service, authentication source, role mapping, enforcement"),
    ("CPPM-002","Security","ClearPass Policy Manager",
     "What is the basic ClearPass authentication flow?",
     "A typical ClearPass flow receives the authentication request, selects the applicable service, validates the identity through the configured authentication source, applies authorization and role mapping logic, and returns the enforcement result to the network device.",
     "1. Identify the network access device.\n2. Identify the authentication method.\n3. Confirm the request matches the intended service.\n4. Confirm the authentication source is available.\n5. Review role mapping and enforcement.\n6. Confirm the network device applies the result.",
     "ClearPass flow, authentication, authorization, service, role mapping, enforcement"),
    ("CPPM-003","Security","ClearPass Role Mapping",
     "What is role mapping used for in ClearPass?",
     "Role mapping assigns one or more roles to an authenticated identity based on configured conditions and attributes. The assigned role can then be used by enforcement logic to determine access.",
     "1. Confirm authentication succeeds.\n2. Review returned identity attributes.\n3. Check role-mapping conditions.\n4. Identify the assigned role.\n5. Confirm the enforcement policy uses that role.",
     "ClearPass role mapping, identity, role, authorization"),
    ("CPPM-004","Security","ClearPass Enforcement",
     "What should I check when ClearPass authenticates a user but gives the wrong VLAN or network access?",
     "Check the selected service, returned identity attributes, role mapping, enforcement policy, and final attributes returned to the network device.",
     "1. Open the affected request.\n2. Confirm the selected service.\n3. Confirm authentication succeeded.\n4. Review assigned roles.\n5. Review enforcement.\n6. Check the attributes returned to the network device.",
     "wrong VLAN, wrong access, ClearPass enforcement, role mapping"),
    ("CPPM-005","Security","ClearPass RADIUS",
     "What should I check when ClearPass does not receive a RADIUS authentication request?",
     "Check reachability between the network access device and ClearPass, RADIUS server configuration, shared secret, configured RADIUS ports, firewall or ACL rules, and whether the network device is sending requests.",
     "1. Verify the ClearPass server address.\n2. Verify the RADIUS shared secret.\n3. Check expected authentication and accounting ports.\n4. Check firewall and ACL rules.\n5. Generate a controlled authentication test.\n6. Confirm whether the request appears in ClearPass request tracking.",
     "RADIUS, ClearPass, no request, shared secret, UDP, firewall"),
    ("LIC-001","Software & Licensing","License Activation",
     "How do I troubleshoot an HPE software license that is not showing?",
     "Verify the product entitlement, account or organization, license identifier, activation or registration state, quantity, and whether the correct portal or workspace is being viewed.",
     "1. Confirm the exact product and license ID.\n2. Confirm the associated account or organization.\n3. Check entitlement status.\n4. Check expected quantity or capacity.\n5. Capture any activation or entitlement error.",
     "license, entitlement, activation, registration, missing license"),
    ("LIC-002","Software & Licensing","License Renewal",
     "What information is needed for an HPE licensing escalation?",
     "Provide the product name, entitlement or license identifier, customer organization, affected account, portal or workspace, exact error, timestamp, expected result, actual result, and supporting evidence.",
     "1. Record the license or entitlement ID.\n2. Capture exact error text.\n3. Record account and organization context.\n4. State expected versus actual behavior.\n5. Attach evidence without exposing secrets.",
     "licensing escalation, entitlement ID, license ID"),
    ("ILO-001","Compute","HPE iLO",
     "What is HPE iLO used for?",
     "HPE Integrated Lights-Out provides remote management capabilities for supported HPE ProLiant servers, including remote hardware monitoring and management functions. Exact capabilities depend on server generation, iLO version, licensing, and configuration.",
     "1. Identify the server generation and iLO version.\n2. Review system health and alerts.\n3. Review management and event logs.\n4. Use supported remote-management functions appropriate to the task.",
     "iLO, remote management, ProLiant, hardware health"),
    ("SRV-001","Compute","HPE ProLiant",
     "What should I check when an HPE ProLiant server does not boot?",
     "Check power state, hardware health indicators, POST behavior, boot-device selection, storage visibility, recent hardware or firmware changes, and system-management logs.",
     "1. Confirm power.\n2. Observe POST or boot messages.\n3. Check hardware health alerts.\n4. Confirm the intended boot device is detected.\n5. Review recent hardware or firmware changes.\n6. Capture the exact error or diagnostic code.",
     "ProLiant boot, POST, boot device, hardware health"),
    ("FW-001","Compute","Firmware Updates",
     "What information should I collect before an HPE server firmware update?",
     "Collect the exact server model and generation, current firmware versions, target firmware version, compatibility requirements, maintenance window, recovery readiness, and rollback or recovery plan.",
     "1. Record current firmware versions.\n2. Confirm the target package applies to the exact model.\n3. Review current applicable HPE documentation.\n4. Confirm the maintenance window.\n5. Confirm recovery planning.",
     "firmware update, compatibility, maintenance window, rollback"),
    ("NET-001","Networking","Aruba CX Switching",
     "What should I check when an Aruba switch port has no connectivity?",
     "Check physical link state, administrative state, port configuration, VLAN assignment, authentication state, interface errors, and the connected endpoint.",
     "1. Check cable and physical link indicators.\n2. Confirm the port is enabled.\n3. Confirm access or trunk configuration.\n4. Confirm the expected VLAN.\n5. Check authentication if applicable.\n6. Review interface error counters.",
     "Aruba switch, CX, port down, VLAN, interface errors"),
    ("NET-002","Networking","VLAN",
     "How should I troubleshoot an Aruba VLAN connectivity problem?",
     "Verify the VLAN exists, the endpoint port assigns or carries the VLAN correctly, trunks carry it where required, the gateway interface exists and is reachable, and routing or ACL policy permits the traffic.",
     "1. Identify source and destination VLANs.\n2. Verify VLAN existence.\n3. Check access-port or trunk VLAN configuration.\n4. Check the gateway or SVI.\n5. Test gateway reachability.\n6. Test the destination.\n7. Check ACL or security policy.",
     "VLAN, Aruba, trunk, SVI, gateway, ACL"),
    ("NET-003","Networking","Wireless",
     "What should I check when an Aruba wireless client cannot connect to an SSID?",
     "Check SSID advertisement, client visibility, authentication method, security credentials, client eligibility, DHCP availability, and whether the issue affects one client or many clients.",
     "1. Confirm the SSID is enabled.\n2. Check whether the client sees it.\n3. Review authentication or security failure details.\n4. Check DHCP and IP assignment.\n5. Compare with a known working client.",
     "Aruba wireless, SSID, client, authentication, DHCP"),
    ("STOR-001","Storage","HPE Alletra",
     "What should I check when a host cannot access an HPE Alletra storage volume?",
     "Check host-to-storage connectivity, initiator and target configuration, zoning or network path, volume presentation or mapping, host multipathing, and storage-side health.",
     "1. Identify host and volume.\n2. Verify host initiator identity.\n3. Verify network or fabric connectivity.\n4. Verify volume mapping or presentation.\n5. Check multipath status.\n6. Check storage alerts and capacity.",
     "Alletra, volume, host connectivity, multipath, mapping"),
    ("STOR-002","Storage","Storage Performance",
     "What should I check when HPE storage performance is slow?",
     "Determine whether the bottleneck is at the host, network or fabric, storage system, volume, workload, or application layer by comparing latency, throughput, I/O rate, path health, and resource utilization.",
     "1. Record the affected workload and time window.\n2. Check host I/O.\n3. Check network or fabric errors.\n4. Check storage latency and throughput.\n5. Compare affected and unaffected workloads.\n6. Correlate with recent changes or unusual load.",
     "storage latency, IOPS, throughput, performance"),
    ("GL-001","Support & Tools","HPE GreenLake",
     "What should I check when a user cannot access an HPE GreenLake service?",
     "Verify the user account, authentication method, organization or workspace membership, assigned permissions, service entitlement, and whether the issue is account-specific or organization-wide.",
     "1. Confirm authentication.\n2. Confirm organization or workspace membership.\n3. Check assigned permissions.\n4. Confirm the service entitlement.\n5. Compare with a known working user when appropriate.",
     "GreenLake, access, workspace, organization, permissions"),
    ("SEC-001","Security","Security",
     "What is the safest approach when troubleshooting a security-sensitive access problem?",
     "Collect the minimum required evidence, avoid exposing credentials or secrets, preserve relevant logs, use least-privilege actions, and follow the approved security and escalation process.",
     "1. Never request passwords or private keys in case notes.\n2. Redact sensitive screenshots.\n3. Preserve timestamps and logs.\n4. Use approved administrative access.\n5. Escalate suspected security incidents.",
     "security, least privilege, credentials, secrets, escalation"),
    ("KB-001","Support & Tools","Knowledge Management",
     "How should an SOP be written so an AI can retrieve an exact answer?",
     "Use one clear question per knowledge record, followed immediately by a direct answer. Add product and topic scope, numbered steps, verification, escalation criteria, and searchable keywords. Keep the record self-contained.",
     "1. Write one question per record.\n2. Put the direct answer first.\n3. Use exact product and feature names.\n4. Keep scope explicit.\n5. Add numbered steps when needed.\n6. Add keywords and metadata.",
     "AI retrieval, RAG, atomic SOP, exact answer, knowledge chunk"),
]


# ============================================================
# DATABASE
# ============================================================
def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    # The old database is deliberately deleted ONE TIME.
    # PDFs are preserved.
    if not RESET_MARKER.exists() and DB_PATH.exists():
        DB_PATH.unlink()

    conn = db()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            filename TEXT,
            family TEXT,
            topic TEXT,
            source_url TEXT,
            content TEXT NOT NULL,
            doc_type TEXT DEFAULT 'PDF',
            created_at TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS kb_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kb_id TEXT UNIQUE NOT NULL,
            family TEXT NOT NULL,
            topic TEXT NOT NULL,
            question TEXT NOT NULL,
            answer TEXT NOT NULL,
            steps TEXT,
            keywords TEXT,
            source TEXT DEFAULT 'Built-in AI Knowledge',
            created_at TEXT NOT NULL
        )
    """)

    if conn.execute("SELECT COUNT(*) FROM kb_records").fetchone()[0] == 0:
        now = datetime.now().isoformat(timespec="seconds")
        conn.executemany("""
            INSERT OR IGNORE INTO kb_records
            (kb_id, family, topic, question, answer, steps, keywords, source, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, [
            (a,b,c,d,e,f,g,"Built-in AI Knowledge",now)
            for a,b,c,d,e,f,g in SEED_KB
        ])

    conn.commit()
    conn.close()

    RESET_MARKER.write_text(
        "New HPE Knowledge Base database initialized.\n",
        encoding="utf-8"
    )


init_db()


# ============================================================
# DATA
# ============================================================
def load_records():
    conn = db()
    rows = conn.execute("""
        SELECT kb_id, family, topic, question, answer, steps, keywords, source
        FROM kb_records ORDER BY id DESC
    """).fetchall()
    conn.close()
    return [dict(x) for x in rows]


def load_documents():
    conn = db()
    rows = conn.execute("""
        SELECT id, title, filename, family, topic, source_url, content, doc_type, created_at
        FROM documents ORDER BY id DESC
    """).fetchall()
    conn.close()
    return [dict(x) for x in rows]


def index_pdf(uploaded_file, family, topic):
    if fitz is None:
        return False, "PyMuPDF is not installed. Add pymupdf to requirements.txt."

    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", uploaded_file.name)
    pdf_path = PDF_DIR / safe
    pdf_path.write_bytes(uploaded_file.getbuffer())

    try:
        pdf = fitz.open(pdf_path)
        text = "\n".join(page.get_text("text") for page in pdf).strip()
        pdf.close()

        if not text:
            return False, "The PDF contains no extractable text."

        conn = db()
        conn.execute("""
            INSERT INTO documents
            (title, filename, family, topic, source_url, content, doc_type, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            Path(uploaded_file.name).stem,
            uploaded_file.name,
            family,
            topic,
            "",
            text,
            "PDF",
            datetime.now().isoformat(timespec="seconds")
        ))
        conn.commit()
        conn.close()
        search.cache_clear()
        return True, f"Indexed {uploaded_file.name}"
    except Exception as e:
        return False, str(e)


# ============================================================
# SEARCH
# ============================================================
@st.cache_data(ttl=30, show_spinner=False)
def search(query, family=None, limit=8):
    records = load_records()
    documents = load_documents()

    if family:
        records = [r for r in records if r["family"] == family]

    q = query.strip().lower()

    if not q:
        return records[:limit], []

    lexical = []
    tokens = re.findall(r"[a-z0-9][a-z0-9\-]+", q)

    for r in records:
        question = r["question"].lower()
        keywords = r["keywords"].lower()
        full = " ".join([
            r["family"], r["topic"], r["question"],
            r["answer"], r["steps"], r["keywords"]
        ]).lower()

        score = 0
        if q in question:
            score += 30
        if q in keywords:
            score += 20
        for t in tokens:
            if t in question:
                score += 6
            elif t in keywords:
                score += 4
            elif t in full:
                score += 1
        if score:
            lexical.append((score, r))

    merged = []

    if TfidfVectorizer and cosine_similarity and records:
        try:
            corpus = [
                " ".join([
                    r["family"], r["topic"], r["question"],
                    r["answer"], r["steps"], r["keywords"]
                ])
                for r in records
            ]
            vec = TfidfVectorizer(stop_words="english", ngram_range=(1,2))
            matrix = vec.fit_transform(corpus)
            qv = vec.transform([query])
            sims = cosine_similarity(qv, matrix)[0]

            for i, sim in enumerate(sims):
                if sim > 0:
                    merged.append((float(sim) * 25, records[i]))
        except Exception:
            pass

    merged.extend(lexical)

    unique = {}
    for score, r in merged:
        unique[r["kb_id"]] = max(score, unique.get(r["kb_id"], 0))

    record_map = {r["kb_id"]: r for r in records}
    ordered = sorted(unique.items(), key=lambda x: x[1], reverse=True)
    result_records = [record_map[k] for k, _ in ordered[:limit]]

    # PDF text search
    doc_scored = []
    for d in documents:
        if family and d["family"] != family:
            continue
        body = (d["content"] or "").lower()
        hits = sum(1 for t in tokens if t in body)
        if q in body:
            hits += 20
        if hits:
            doc_scored.append((hits, d))

    doc_scored.sort(key=lambda x: x[0], reverse=True)
    return result_records, [d for _, d in doc_scored[:limit]]


# ============================================================
# STYLE
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
  --navy:#061d2b;
  --navy2:#073247;
  --ink:#0b3154;
  --muted:#637b8e;
  --green:#00bfa5;
  --cyan:#00dfcf;
  --border:#d8e7ec;
}

* { box-sizing:border-box; }
html, body, [class*="css"] { font-family:Inter,Arial,sans-serif !important; }

.stApp {
  background:
    radial-gradient(circle at 10% 35%, rgba(0,207,195,.16), transparent 23%),
    radial-gradient(circle at 90% 45%, rgba(0,190,177,.12), transparent 22%),
    linear-gradient(180deg,#eaf8f7 0%,#f6fbfc 58%,#edf8f7 100%);
}

header[data-testid="stHeader"] { display:none; }

.block-container {
  max-width:1536px !important;
  padding:0 22px 18px !important;
}

div[data-testid="stToolbar"] { display:none !important; }

.hero {
  height:245px;
  margin:0 -22px 10px;
  padding:16px 38px 12px;
  color:white;
  position:relative;
  overflow:hidden;
  background:
    radial-gradient(ellipse at 78% 47%,rgba(0,239,218,.23),transparent 23%),
    radial-gradient(ellipse at 31% 12%,rgba(0,119,158,.30),transparent 31%),
    linear-gradient(117deg,#041d2a 0%,#062a3a 52%,#062b38 100%);
}

.hero:before {
  content:"";
  position:absolute;
  left:165px; top:-115px;
  width:850px; height:270px;
  border:1px solid rgba(0,230,220,.25);
  border-radius:50%;
  transform:rotate(-8deg);
  box-shadow:
    0 0 0 40px rgba(0,230,220,.025),
    0 0 0 80px rgba(0,230,220,.018);
}

.hero:after {
  content:"";
  position:absolute;
  right:-30px; top:0;
  width:500px; height:245px;
  background:
    linear-gradient(112deg,transparent 0 43%,rgba(0,255,230,.12) 43.5% 44%,transparent 44.5%),
    linear-gradient(70deg,transparent 0 56%,rgba(0,255,230,.09) 56.5% 57%,transparent 57.5%);
}

.hero-grid {
  position:relative;
  z-index:2;
  display:grid;
  grid-template-columns:205px 1fr 155px;
  height:100%;
  align-items:start;
}

.brand { padding-top:2px; }

.hpe-logo {
  width:80px;
  height:25px;
  border:5px solid #00d9bd;
  border-bottom:0;
  margin-bottom:2px;
}

.hpe-text {
  font-size:30px;
  line-height:29px;
  font-weight:800;
  letter-spacing:-1.5px;
}

.kb-text {
  font-size:17px;
  line-height:18px;
  font-weight:800;
}

.brand-rule {
  width:174px;
  height:1px;
  background:rgba(255,255,255,.48);
  margin:15px 0 9px;
}

.brand-copy {
  font-size:10px;
  line-height:1.45;
  color:#d9edf0;
}

.hero-center { text-align:center; }

.hero-title {
  font-size:43px;
  line-height:47px;
  font-weight:800;
  letter-spacing:-1.8px;
  margin:2px 0 3px;
}

.hero-title span { color:#00e5c6; }

.hero-sub {
  font-size:15px;
  color:#d7eef1;
  margin-bottom:13px;
}

.hero-search {
  max-width:820px;
  margin:auto;
}

div[data-testid="stTextInput"] { margin:0 !important; }
div[data-testid="stTextInput"] label { display:none !important; }
div[data-testid="stTextInput"] input {
  height:52px !important;
  border-radius:28px !important;
  border:2px solid rgba(0,205,190,.38) !important;
  background:#fff !important;
  color:#36526a !important;
  font-size:14px !important;
  padding:0 48px 0 42px !important;
  box-shadow:0 7px 20px rgba(0,30,50,.15) !important;
}

.hero-search-icon {
  position:absolute;
  z-index:5;
  margin:15px 0 0 17px;
  color:#153b59;
  font-size:22px;
}

.try {
  margin-top:9px;
  display:flex;
  justify-content:center;
  align-items:center;
  gap:7px;
  color:#e3f5f5;
  font-size:9px;
}

.try-chip {
  padding:6px 12px;
  border:1px solid rgba(255,255,255,.32);
  border-radius:18px;
  background:rgba(255,255,255,.10);
  color:white;
}

.hero-right {
  text-align:right;
  padding-top:60px;
  font-size:15px;
  line-height:1.05;
  font-weight:700;
}

.hero-right .next {
  color:#e2f8f5;
  font-size:10px;
  font-weight:500;
  display:block;
  margin-top:2px;
}

.hero-right .dash {
  color:#00e6c7;
  font-size:25px;
  line-height:20px;
}

.family-row {
  display:grid;
  grid-template-columns:repeat(6,1fr);
  gap:9px;
  margin:0 0 11px;
}

.family-card {
  height:137px;
  background:rgba(255,255,255,.92);
  border:1px solid rgba(214,228,234,.95);
  border-radius:15px;
  padding:14px 13px 10px;
  box-shadow:0 5px 18px rgba(36,76,92,.07);
  position:relative;
}

.family-icon {
  width:38px;
  height:38px;
  border-radius:10px;
  display:flex;
  align-items:center;
  justify-content:center;
  color:#fff;
  font-size:21px;
  font-weight:700;
  margin-bottom:7px;
}

.mint{background:linear-gradient(145deg,#08bda0,#10a7bf)}
.blue{background:linear-gradient(145deg,#1476e7,#3768eb)}
.violet{background:linear-gradient(145deg,#8157e8,#a15ce9)}
.orange{background:linear-gradient(145deg,#ff9b40,#ff7a19)}
.red{background:linear-gradient(145deg,#ff6570,#f13f4b)}
.slate{background:linear-gradient(145deg,#607a92,#405e78)}

.family-name {
  color:#0a3154;
  font-size:16px;
  line-height:18px;
  font-weight:800;
}

.family-desc {
  color:#536f82;
  font-size:10px;
  line-height:13px;
  margin-top:4px;
  padding-right:19px;
}

.family-arrow {
  position:absolute;
  right:12px;
  bottom:14px;
  width:29px;
  height:29px;
  border-radius:50%;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:17px;
  font-weight:700;
}

.arrow-mint{background:#d3f7ef;color:#00a88e}
.arrow-blue{background:#dcecff;color:#2a74df}
.arrow-violet{background:#ece4ff;color:#7c54db}
.arrow-orange{background:#fff0df;color:#f17c1d}
.arrow-red{background:#ffe3e6;color:#f24f5a}
.arrow-slate{background:#e8eef3;color:#516c83}

.family-click {
  margin-top:-42px !important;
  position:relative;
  z-index:10;
}

.family-click div.stButton > button {
  height:44px !important;
  min-height:44px !important;
  opacity:0 !important;
  cursor:pointer !important;
  border:0 !important;
  background:transparent !important;
}

.main-grid {
  display:grid;
  grid-template-columns:1.02fr .98fr;
  gap:12px;
}

.panel {
  background:rgba(255,255,255,.90);
  border:1px solid rgba(215,229,234,.96);
  border-radius:18px;
  box-shadow:0 6px 19px rgba(35,74,91,.07);
  padding:16px;
}

.ai-panel { min-height:405px; }
.docs-panel { min-height:255px; }
.topics-panel { min-height:139px; }

.ai-head {
  display:flex;
  align-items:center;
  gap:10px;
}

.robot {
  width:48px; height:48px;
  border-radius:50%;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:24px;
  background:linear-gradient(145deg,#e2faf6,#c9f4ee);
  border:1px solid #c0ebe4;
}

.panel-title {
  color:#0a3154;
  font-size:17px;
  font-weight:800;
}

.panel-sub {
  color:#718596;
  font-size:10px;
  margin-top:1px;
}

.beta {
  display:inline-block;
  margin-left:4px;
  padding:3px 7px;
  border-radius:8px;
  background:#dff6ff;
  color:#087da1;
  font-size:8px;
  font-weight:800;
}

.powered {
  margin-left:auto;
  padding:8px 12px;
  border-radius:18px;
  background:#f0faf9;
  border:1px solid #d8eee9;
  color:#008e79;
  font-size:9px;
}

.user-msg {
  margin:13px 0 7px 12px;
  padding:9px 12px;
  border-radius:9px;
  background:#eef3f6;
  color:#243f55;
  font-size:10px;
}

.ai-answer {
  margin-left:12px;
  padding:12px 13px;
  border-radius:0 10px 10px 10px;
  background:#f1f6f8;
  border-left:3px solid #00bfa5;
  color:#263f54;
  font-size:10px;
  line-height:1.45;
}

.ai-step {
  display:flex;
  gap:9px;
  margin:6px 0;
}

.ai-num {
  flex:0 0 20px;
  width:20px;height:20px;
  border-radius:50%;
  background:#09b89f;
  color:white;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:9px;
  font-weight:800;
}

.ai-doc {
  margin-top:9px;
  padding:9px;
  border-radius:7px;
  background:white;
  border:1px solid #dce7eb;
  display:flex;
  align-items:center;
  gap:8px;
}

.pdf-icon {
  width:28px;height:31px;
  border-radius:4px;
  background:#ee4c45;
  color:white;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:13px;
}

.ai-actions {
  display:flex;
  gap:7px;
  margin:9px 12px;
}

.ai-actions span {
  border:1px solid #d6e4e8;
  background:white;
  padding:6px 10px;
  border-radius:16px;
  color:#466274;
  font-size:8px;
}

.follow {
  margin:7px 0 0;
  border:1px solid #d5e6ea;
  background:white;
  border-radius:15px;
  height:36px;
  display:flex;
  align-items:center;
  padding:0 12px;
  color:#728696;
  font-size:9px;
}

.follow-arrow {
  margin-left:auto;
  width:25px;height:25px;
  background:#08bda4;
  color:white;
  border-radius:50%;
  display:flex;
  justify-content:center;
  align-items:center;
  font-size:15px;
}

.section-head {
  display:flex;
  align-items:flex-start;
  justify-content:space-between;
}

.section-head-title {
  display:flex;
  align-items:center;
  gap:9px;
}

.section-emoji { font-size:23px; }

.view-all {
  color:#0069db;
  font-size:10px;
  font-weight:700;
}

.section-caption {
  color:#6f8595;
  font-size:9px;
  margin:0 0 10px 32px;
}

.doc-grid {
  display:grid;
  grid-template-columns:repeat(4,1fr);
  gap:9px;
}

.doc-card {
  border:1px solid #dbe7ec;
  border-radius:9px;
  background:white;
  padding:7px;
  min-height:178px;
}

.doc-cover {
  height:72px;
  border-radius:5px;
  display:flex;
  align-items:flex-end;
  padding:7px;
  color:white;
  font-size:8px;
  line-height:9px;
  font-weight:800;
  background:
    radial-gradient(circle at 72% 20%,rgba(0,241,222,.45),transparent 24%),
    linear-gradient(140deg,#061d2b,#0a4352 65%,#00a58c);
}

.doc-cover:nth-child(2) { background:linear-gradient(140deg,#071f2f,#173e63,#0bbca1); }

.doc-title {
  color:#153753;
  font-size:9px;
  line-height:11px;
  font-weight:800;
  margin-top:7px;
  min-height:34px;
}

.doc-meta {
  color:#8999a4;
  font-size:8px;
  margin-top:4px;
}

.topic-grid {
  display:grid;
  grid-template-columns:repeat(3,1fr);
  gap:8px;
}

.topic-card {
  min-height:49px;
  border:1px solid #dbe7ec;
  background:white;
  border-radius:9px;
  padding:8px 9px;
  display:flex;
  align-items:center;
  gap:8px;
}

.topic-icon {
  width:27px;height:27px;
  border-radius:50%;
  background:#e6f7f4;
  color:#00a78f;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:14px;
}

.topic-name {
  color:#173a56;
  font-size:9px;
  font-weight:700;
}

.topic-count {
  color:#91a0aa;
  font-size:7px;
  margin-top:1px;
}

.bottom-strip {
  height:61px;
  margin-top:11px;
  border:1px solid #dce8ec;
  border-radius:16px;
  background:rgba(255,255,255,.92);
  display:grid;
  grid-template-columns:repeat(5,1fr);
  align-items:center;
}

.stat {
  display:flex;
  align-items:center;
  justify-content:center;
  gap:10px;
  border-right:1px solid #dce8ec;
}

.stat:last-child { border-right:0; }

.stat-icon {
  font-size:21px;
  color:#00a991;
}

.stat-value {
  color:#173852;
  font-size:10px;
  font-weight:800;
}

.stat-label {
  color:#788c9a;
  font-size:8px;
  margin-top:2px;
}

.search-result {
  background:white;
  border:1px solid #dbe7ec;
  border-radius:12px;
  padding:13px;
  margin:8px 0;
}

.result-id {
  color:#7c909d;
  font-size:8px;
}

.result-q {
  color:#0c3455;
  font-size:12px;
  font-weight:800;
  margin-top:4px;
}

.result-a {
  color:#536c7e;
  font-size:10px;
  line-height:1.45;
  margin-top:5px;
}

.family-banner {
  margin:4px 0 12px;
  padding:18px 22px;
  border-radius:17px;
  color:white;
  background:linear-gradient(120deg,#061d2b,#074356);
}

.family-banner h1 { margin:0; font-size:27px; }
.family-banner p { margin:4px 0 0; color:#cfe9eb; font-size:11px; }

.topic-directory {
  background:white;
  border:1px solid #dbe7ec;
  border-radius:12px;
  padding:12px;
  margin-bottom:9px;
}

.source-card {
  background:white;
  border:1px solid #dbe7ec;
  border-radius:11px;
  padding:11px;
  margin-bottom:8px;
}

.source-card a {
  color:#008f7a;
  text-decoration:none;
  font-weight:700;
  font-size:10px;
}

.footer {
  text-align:center;
  color:#7d909c;
  font-size:8px;
  padding:10px 0 0;
}

div.stButton > button {
  border-radius:18px !important;
  border:1px solid #cfe4e8 !important;
  background:white !important;
  color:#008f7b !important;
  font-size:9px !important;
  font-weight:700 !important;
  min-height:31px !important;
}

.family-click div.stButton > button {
  opacity:0 !important;
  height:48px !important;
  min-height:48px !important;
  margin-top:-45px !important;
  position:relative !important;
  z-index:30 !important;
}

.small-button div.stButton > button {
  min-height:26px !important;
  font-size:8px !important;
  padding:2px 8px !important;
}

@media (max-width: 1100px) {
  .hero-grid { grid-template-columns:160px 1fr 120px; }
  .hero-title { font-size:34px; }
  .family-card { height:145px; }
  .main-grid { grid-template-columns:1fr; }
}

@media (max-width: 800px) {
  .hero { height:auto; min-height:350px; }
  .hero-grid { grid-template-columns:1fr; }
  .hero-right { display:none; }
  .brand { text-align:center; }
  .brand-rule { margin-left:auto;margin-right:auto; }
  .hero-title { margin-top:10px; }
  .family-row { grid-template-columns:repeat(2,1fr); }
  .doc-grid { grid-template-columns:repeat(2,1fr); }
  .topic-grid { grid-template-columns:1fr; }
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# UI HELPERS
# ============================================================
def eh(value):
    return escape(str(value or ""))


def render_hero():
    st.markdown("""
    <div class="hero">
      <div class="hero-grid">
        <div class="brand">
          <div class="hpe-logo"></div>
          <div class="hpe-text">HPE</div>
          <div class="kb-text">Knowledge Base</div>
          <div class="brand-rule"></div>
          <div class="brand-copy">Find answers. Solve faster.<br>Power your support with HPE knowledge.</div>
        </div>

        <div class="hero-center">
          <div class="hero-title">How can we help you <span>today?</span></div>
          <div class="hero-sub">Search HPE documentation, guides, and solutions with AI.</div>
          <div class="hero-search">
            <div class="hero-search-icon">⌕</div>
    """, unsafe_allow_html=True)

    st.text_input(
        "Global search",
        placeholder="Ask a question or search for a document...",
        key="global_query",
        label_visibility="collapsed",
    )

    st.markdown("""
          </div>
          <div class="try">
            <b>Try asking:</b>
            <span class="try-chip">How to renew a license?</span>
            <span class="try-chip">ClearPass troubleshooting</span>
            <span class="try-chip">iLO configuration</span>
            <span class="try-chip">Aruba switch setup</span>
            <span class="try-chip">Gen11 firmware update</span>
          </div>
        </div>

        <div class="hero-right">
          Accelerating<br>what's next<br><span class="next">together</span>
          <div class="dash">—</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)


def open_group(group):
    st.session_state.view = "group"
    st.session_state.selected_group = group
    st.session_state.selected_topic = None
    st.rerun()


def render_family_cards():
    names = list(PRODUCT_GROUPS.keys())
    cols = st.columns(6, gap="small")

    for i, group in enumerate(names):
        data = PRODUCT_GROUPS[group]
        with cols[i]:
            st.markdown(
                f"""
                <div class="family-card">
                  <div class="family-icon {data['class']}">{data['icon']}</div>
                  <div class="family-name">{eh(group)}</div>
                  <div class="family-desc">{eh(data['description'])}</div>
                  <div class="family-arrow arrow-{data['class']}">→</div>
                </div>
                """,
                unsafe_allow_html=True
            )
            st.markdown('<div class="family-click">', unsafe_allow_html=True)
            if st.button("Open family", key=f"family_{group}", use_container_width=True):
                open_group(group)
            st.markdown("</div>", unsafe_allow_html=True)


def answer_card(r):
    st.markdown(
        f"""
        <div class="search-result">
          <div class="result-id">{eh(r['kb_id'])} • {eh(r['family'])} • {eh(r['topic'])}</div>
          <div class="result-q">{eh(r['question'])}</div>
          <div class="result-a">{eh(r['answer'])}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    with st.expander("Show troubleshooting steps"):
        for line in (r["steps"] or "").splitlines():
            st.markdown(eh(line))
        st.caption("Search terms: " + r["keywords"])


def render_ai_answer(query, family=None):
    records, docs = search(query, family=family, limit=6)

    if not records and not docs:
        st.warning("No matching knowledge was found. Try the product name, feature, model, version, or exact error message.")
        return

    best = records[0] if records else None

    if best:
        st.markdown(
            f"""
            <div class="panel ai-panel">
              <div class="ai-head">
                <div class="robot">🤖</div>
                <div>
                  <div class="panel-title">HPE AI Assistant <span class="beta">BETA</span></div>
                  <div class="panel-sub">Get instant answers from HPE documentation.</div>
                </div>
                <div class="powered">✦ Powered by HPE Knowledge</div>
              </div>

              <div class="user-msg">♙ &nbsp; {eh(query)}</div>

              <div class="ai-answer">
                <b>Here is the most relevant answer from HPE Knowledge:</b><br><br>
                {eh(best['answer'])}
              </div>
            """,
            unsafe_allow_html=True
        )

        # Render exact steps like the screenshot.
        steps = (best["steps"] or "").splitlines()
        for i, line in enumerate(steps, 1):
            clean = re.sub(r"^\s*\d+\.\s*", "", line)
            st.markdown(
                f"""
                <div style="margin-left:25px;margin-top:-2px;">
                  <div class="ai-step">
                    <div class="ai-num">{i}</div>
                    <div style="font-size:9px;color:#324e63;padding-top:3px;">{eh(clean)}</div>
                  </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.markdown(
            f"""
              <div class="ai-doc">
                <div class="pdf-icon">▤</div>
                <div style="flex:1;">
                  <div style="font-size:9px;font-weight:800;color:#173a56;">
                    {eh(best['topic'])} — HPE Knowledge Article
                  </div>
                  <div style="font-size:8px;color:#8a9aa5;">AI Knowledge • {eh(best['kb_id'])}</div>
                </div>
                <div style="font-size:16px;color:#173a56;">↗</div>
              </div>

              <div class="ai-actions">
                <span>✦ Summarize this document</span>
                <span>⌕ Show troubleshooting steps</span>
                <span>▤ List related documents</span>
              </div>

              <div class="follow">
                ♧ &nbsp; Ask a follow-up question...
                <span class="follow-arrow">→</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    if len(records) > 1:
        st.markdown("### Related answers")
        for r in records[1:]:
            answer_card(r)

    if docs:
        st.markdown("### Related documents")
        render_documents(docs)


def render_documents(docs):
    cols = st.columns(min(4, max(1, len(docs))))
    for i, d in enumerate(docs[:4]):
        with cols[i % len(cols)]:
            st.markdown(
                f"""
                <div class="doc-card">
                  <div class="doc-cover">HPE<br>KNOWLEDGE</div>
                  <div class="doc-title">{eh(d['title'])}</div>
                  <div class="doc-meta">⌁ PDF • Indexed</div>
                </div>
                """,
                unsafe_allow_html=True
            )
            if st.button("View", key=f"doc_{d['id']}", use_container_width=True):
                st.session_state.view = "document"
                st.session_state.selected_document = d["id"]
                st.rerun()


def render_featured():
    docs = load_documents()

    st.markdown("""
    <div class="panel docs-panel">
      <div class="section-head">
        <div>
          <div class="section-head-title">
            <span class="section-emoji">⭐</span>
            <span class="panel-title">Featured Documents</span>
          </div>
          <div class="section-caption">Most useful and commonly accessed resources.</div>
        </div>
        <div class="view-all">View All →</div>
      </div>
    """, unsafe_allow_html=True)

    if docs:
        render_documents(docs[:4])
    else:
        featured = [
            ("ARBSOP040 - HPE ClearPass Licensing Server Overview", "ClearPass"),
            ("HPE Gen11 Server Setup and Configuration Guide", "Compute"),
            ("Aruba Switch Configuration Guide", "Networking"),
            ("HPE iLO 6 User Guide", "iLO"),
        ]
        cols = st.columns(4)
        for i, (title, label) in enumerate(featured):
            with cols[i]:
                st.markdown(
                    f"""
                    <div class="doc-card">
                      <div class="doc-cover">{eh(label)}<br>GUIDE</div>
                      <div class="doc-title">{eh(title)}</div>
                      <div class="doc-meta">⌁ PDF • Knowledge source</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                if st.button("View", key=f"feature_{i}", use_container_width=True):
                    st.session_state.global_query = title
                    st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def render_popular_topics():
    topics = [
        ("License Activation", "Software & Licensing", "⌕", "125+ documents"),
        ("ClearPass Troubleshooting", "Security", "⌘", "98+ documents"),
        ("iLO Configuration", "Compute", "▣", "110+ documents"),
        ("Firmware Updates", "Compute", "⟳", "86+ documents"),
        ("Switch Setup", "Networking", "⌘", "102+ documents"),
        ("Alletra Storage", "Storage", "▤", "74+ documents"),
    ]

    st.markdown("""
    <div class="panel topics-panel">
      <div class="section-head">
        <div>
          <div class="section-head-title">
            <span class="section-emoji">🔥</span>
            <span class="panel-title">Popular Topics</span>
          </div>
          <div class="section-caption">Frequently searched by HPE support teams.</div>
        </div>
        <div class="view-all">View All →</div>
      </div>
    """, unsafe_allow_html=True)

    cols = st.columns(3, gap="small")
    for i, (topic, group, icon, count) in enumerate(topics):
        with cols[i % 3]:
            st.markdown(
                f"""
                <div class="topic-card">
                  <div class="topic-icon">{icon}</div>
                  <div style="flex:1">
                    <div class="topic-name">{eh(topic)}</div>
                    <div class="topic-count">{eh(count)}</div>
                  </div>
                  <div style="color:#6b8191;font-size:12px;">›</div>
                </div>
                """,
                unsafe_allow_html=True
            )
            if st.button("Open", key=f"pop_{i}", use_container_width=True):
                st.session_state.view = "group"
                st.session_state.selected_group = group
                st.session_state.selected_topic = topic
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def render_bottom_strip():
    stats = [
        ("♧", "10,000+", "Technical Documents"),
        ("✦", "AI-Powered", "Instant Answers"),
        ("◷", "Up-to-Date", "Latest HPE Resources"),
        ("♧", "Easy to Navigate", "Find What You Need"),
        ("♢", "Trusted Content", "From HPE Experts"),
    ]

    html = '<div class="bottom-strip">'
    for icon, value, label in stats:
        html += (
            f'<div class="stat">'
            f'<div class="stat-icon">{icon}</div>'
            f'<div><div class="stat-value">{eh(value)}</div>'
            f'<div class="stat-label">{eh(label)}</div></div>'
            f'</div>'
        )
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# HOME
# ============================================================
def render_home():
    render_hero()

    query = st.session_state.get("global_query", "").strip()

    # Search results appear directly below the hero, before the six tiles.
    if query:
        st.markdown('<div class="panel" style="margin-bottom:11px;">', unsafe_allow_html=True)
        render_ai_answer(query)
        st.markdown("</div>", unsafe_allow_html=True)

    render_family_cards()

    left, right = st.columns([1.02, .98], gap="small")

    with left:
        # Exact left-side AI panel proportions.
        if not query:
            st.markdown("""
            <div class="panel ai-panel">
              <div class="ai-head">
                <div class="robot">🤖</div>
                <div>
                  <div class="panel-title">HPE AI Assistant <span class="beta">BETA</span></div>
                  <div class="panel-sub">Get instant answers from HPE documentation.</div>
                </div>
                <div class="powered">✦ Powered by HPE Knowledge</div>
              </div>

              <div class="user-msg">♙ &nbsp; How do I troubleshoot ClearPass licensing issues?</div>

              <div class="ai-answer">
                <b>Here are the steps to troubleshoot ClearPass licensing issues based on HPE documentation:</b>
              </div>

              <div style="margin:8px 0 0 25px;">
                <div class="ai-step"><div class="ai-num">1</div><div style="font-size:9px;color:#324e63;padding-top:3px;">Verify license status on the ClearPass Administration Portal.</div></div>
                <div class="ai-step"><div class="ai-num">2</div><div style="font-size:9px;color:#324e63;padding-top:3px;">Check the license server connectivity.</div></div>
                <div class="ai-step"><div class="ai-num">3</div><div style="font-size:9px;color:#324e63;padding-top:3px;">Ensure the correct license file is installed.</div></div>
                <div class="ai-step"><div class="ai-num">4</div><div style="font-size:9px;color:#324e63;padding-top:3px;">Review logs for licensing errors.</div></div>
                <div class="ai-step"><div class="ai-num">5</div><div style="font-size:9px;color:#324e63;padding-top:3px;">If the issue persists, refer to the relevant ClearPass licensing article.</div></div>
              </div>

              <div class="ai-doc">
                <div class="pdf-icon">▤</div>
                <div style="flex:1;">
                  <div style="font-size:9px;font-weight:800;color:#173a56;">ARBSOP040 - HPE ClearPass Licensing Server Overview</div>
                  <div style="font-size:8px;color:#8a9aa5;">⌁ PDF • 2.4 MB • Knowledge Source</div>
                </div>
                <div style="font-size:16px;color:#173a56;">↗</div>
              </div>

              <div class="ai-actions">
                <span>✦ Summarize this document</span>
                <span>⌕ Show troubleshooting steps</span>
                <span>▤ List related documents</span>
              </div>

              <div class="follow">
                ♧ &nbsp; Ask a follow-up question...
                <span class="follow-arrow">→</span>
              </div>
            </div>
            """, unsafe_allow_html=True)

            if st.button("Ask this question", key="demo_question", use_container_width=False):
                st.session_state.global_query = "How do I troubleshoot ClearPass licensing issues?"
                st.rerun()
        else:
            # Search already rendered above; leave the visual position clean.
            st.markdown("""
            <div class="panel ai-panel">
              <div class="ai-head">
                <div class="robot">🤖</div>
                <div>
                  <div class="panel-title">HPE AI Assistant <span class="beta">BETA</span></div>
                  <div class="panel-sub">Ask another question from the search bar above.</div>
                </div>
              </div>
            </div>
            """, unsafe_allow_html=True)

    with right:
        render_featured()
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        render_popular_topics()

    render_bottom_strip()


# ============================================================
# PRODUCT GROUP PAGE
# ============================================================
def render_group():
    group = st.session_state.get("selected_group")
    if group not in PRODUCT_GROUPS:
        st.session_state.view = "home"
        st.rerun()

    data = PRODUCT_GROUPS[group]

    if st.button("← Back to Knowledge Base", key="back_group"):
        st.session_state.view = "home"
        st.rerun()

    st.markdown(
        f"""
        <div class="family-banner">
          <h1>{eh(data['icon'])} {eh(group)}</h1>
          <p>{eh(data['description'])}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    # If a topic was selected from Popular Topics, focus it.
    selected_topic = st.session_state.get("selected_topic")

    st.markdown(
        f"""
        <div class="panel" style="margin-bottom:10px;">
          <div class="panel-title">Product Families</div>
          <div class="panel-sub">All major HPE and Aruba products mapped to this product group.</div>
          <div style="height:8px"></div>
          <div class="topic-grid">
        """,
        unsafe_allow_html=True
    )

    for product in data["products"]:
        st.markdown(
            f"""
            <div class="topic-directory">
              <div style="font-size:10px;font-weight:800;color:#153954;">{eh(product)}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    st.markdown("</div></div>", unsafe_allow_html=True)

    st.markdown("""
    <div class="panel" style="margin-bottom:10px;">
      <div class="panel-title">Topics & Solutions</div>
      <div class="panel-sub">Click a topic to view its related AI answers and sources.</div>
      <div style="height:8px"></div>
    """, unsafe_allow_html=True)

    topic_cols = st.columns(3, gap="small")
    for i, topic in enumerate(data["topics"]):
        with topic_cols[i % 3]:
            st.markdown(
                f"""
                <div class="topic-card">
                  <div class="topic-icon">{eh(data['icon'])}</div>
                  <div style="flex:1">
                    <div class="topic-name">{eh(topic)}</div>
                    <div class="topic-count">View related knowledge</div>
                  </div>
                  <div style="color:#6b8191;">›</div>
                </div>
                """,
                unsafe_allow_html=True
            )
            if st.button("Open", key=f"group_topic_{group}_{i}", use_container_width=True):
                st.session_state.selected_topic = topic
                st.session_state.view = "topic"
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("""
    <div class="panel">
      <div class="panel-title">Source Library</div>
      <div class="panel-sub">Official/public documentation sources associated with this product group.</div>
      <div style="height:8px"></div>
    """, unsafe_allow_html=True)

    source_cols = st.columns(4, gap="small")
    for i, (name, url) in enumerate(data["sources"]):
        with source_cols[i % 4]:
            st.markdown(
                f"""
                <div class="source-card">
                  <div style="font-size:10px;font-weight:800;color:#173a56;">{eh(name)}</div>
                  <div class="small-muted">Official/public source</div>
                  <div style="height:7px"></div>
                  <a href="{eh(url)}" target="_blank">Open source →</a>
                </div>
                """,
                unsafe_allow_html=True
            )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("### Ask AI about this product group")
    q = st.text_input(
        "Family search",
        placeholder=f"Ask a question about {group}...",
        key=f"family_question_{group}",
        label_visibility="collapsed"
    )
    if q:
        render_ai_answer(q, family=group)


# ============================================================
# TOPIC PAGE
# ============================================================
def render_topic():
    group = st.session_state.get("selected_group")
    topic = st.session_state.get("selected_topic")

    if group not in PRODUCT_GROUPS:
        st.session_state.view = "home"
        st.rerun()

    if st.button("← Back to product family", key="back_topic"):
        st.session_state.view = "group"
        st.rerun()

    st.markdown(
        f"""
        <div class="family-banner">
          <h1>{eh(topic)}</h1>
          <p>{eh(group)} • Related topics, AI answers and indexed documents</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Search the built-in atomic records using topic + family.
    records = [
        r for r in load_records()
        if r["family"] == group and (
            topic.lower() in r["topic"].lower()
            or any(word in (r["question"] + " " + r["keywords"]).lower()
                   for word in topic.lower().split() if len(word) > 3)
        )
    ]

    if records:
        for r in records:
            answer_card(r)
    else:
        st.info("There is currently no atomic AI record mapped directly to this topic.")

    st.markdown("### Search this topic")
    q = st.text_input(
        "Topic search",
        placeholder=f"Ask a specific {topic} question...",
        key=f"topic_question_{group}_{topic}",
        label_visibility="collapsed"
    )
    if q:
        render_ai_answer(q, family=group)


# ============================================================
# DOCUMENT PAGE
# ============================================================
def render_document():
    doc_id = st.session_state.get("selected_document")
    doc = next((x for x in load_documents() if x["id"] == doc_id), None)

    if not doc:
        st.error("Document not found.")
        if st.button("Return"):
            st.session_state.view = "home"
            st.rerun()
        return

    if st.button("← Back", key="back_document"):
        st.session_state.view = "home"
        st.rerun()

    st.markdown(
        f"""
        <div class="family-banner">
          <h1>{eh(doc['title'])}</h1>
          <p>{eh(doc['family'])} • {eh(doc['topic'])} • {eh(doc['filename'])}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    find = st.text_input(
        "Find in document",
        placeholder="Search within this PDF...",
        label_visibility="collapsed"
    )

    content = doc["content"] or ""
    if find:
        pieces = re.split(r"(?<=[.!?])\s+", content)
        hits = [x for x in pieces if find.lower() in x.lower()]
        if hits:
            for h in hits[:50]:
                st.markdown(f'<div class="search-result">{eh(h)}</div>', unsafe_allow_html=True)
        else:
            st.info("No matching text found.")
    else:
        st.text_area("Document text", content, height=620, label_visibility="collapsed")


# ============================================================
# ADMIN
# Hidden from the visual reference on the home page.
# Open by adding ?admin=1 to the app URL.
# ============================================================
def render_admin():
    if st.button("← Back to Knowledge Base", key="admin_back"):
        st.session_state.view = "home"
        st.rerun()

    st.markdown("""
    <div class="family-banner">
      <h1>Knowledge Management</h1>
      <p>Upload SOPs and PDFs into the HPE Knowledge Base.</p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        group = st.selectbox("Product Group", list(PRODUCT_GROUPS.keys()))
    with c2:
        topic = st.text_input("Topic", placeholder="e.g. ClearPass Licensing")

    file = st.file_uploader("Upload PDF", type=["pdf"])

    if file and st.button("Index PDF", type="primary"):
        ok, msg = index_pdf(file, group, topic or "General")
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    st.markdown("### Indexed documents")
    docs = load_documents()
    if not docs:
        st.info("No PDFs indexed yet.")
    else:
        for d in docs:
            st.markdown(
                f"""
                <div class="search-result">
                  <div class="result-q">{eh(d['title'])}</div>
                  <div class="result-a">{eh(d['family'])} • {eh(d['topic'])} • {eh(d['created_at'])}</div>
                </div>
                """,
                unsafe_allow_html=True
            )


# ============================================================
# ROUTER
# ============================================================
for key, default in [
    ("view","home"),
    ("selected_group",None),
    ("selected_topic",None),
    ("selected_document",None),
    ("global_query",""),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# Admin is intentionally absent from the screenshot UI.
# Use ?admin=1 to enter the uploader without adding an extra
# visible button that would break the reference layout.
try:
    if str(st.query_params.get("admin","")).lower() in {"1","true","yes"}:
        st.session_state.view = "admin"
except Exception:
    pass

if st.session_state.view == "home":
    render_home()
elif st.session_state.view == "group":
    render_group()
elif st.session_state.view == "topic":
    render_topic()
elif st.session_state.view == "document":
    render_document()
elif st.session_state.view == "admin":
    render_admin()
else:
    st.session_state.view = "home"
    st.rerun()

st.markdown(
    '<div class="footer">HPE Knowledge Base • AI-assisted technical knowledge search • '
    'Validate production procedures against current authoritative HPE documentation.</div>',
    unsafe_allow_html=True
)
