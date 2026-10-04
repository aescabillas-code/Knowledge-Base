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
  position:relative !important;
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

.hero-search { max-width:820px; margin:auto; }

.st-key-hero_search_area { position:relative; z-index:20; margin:-102px auto 8px !important; max-width:830px; }
.st-key-hero_search_area form { background:transparent !important; border:0 !important; padding:0 !important; }
.st-key-hero_search_area div[data-testid="stTextInput"] { margin:0 !important; }
.st-key-hero_search_area div[data-testid="stTextInput"] label { display:none !important; }
.st-key-hero_search_area div[data-testid="stTextInput"] input { height:54px !important; border-radius:29px !important; border:2px solid rgba(0,205,190,.40) !important; background:#fff !important; color:#36526a !important; font-size:14px !important; padding:0 20px 0 22px !important; box-shadow:0 7px 20px rgba(0,30,50,.20) !important; }
.st-key-hero_search_area div[data-testid="stFormSubmitButton"] button { width:52px !important; height:52px !important; min-height:52px !important; padding:0 !important; margin-top:0 !important; border-radius:50% !important; border:0 !important; background:#08bca3 !important; color:white !important; box-shadow:0 4px 12px rgba(0,160,140,.25) !important; font-size:25px !important; line-height:1 !important; }
.st-key-hero_search_area .try-label { text-align:left; display:inline-block; color:#e5f6f5; font-size:10px; font-weight:600; margin:6px 0 0; text-shadow:0 1px 3px rgba(0,0,0,.25); }
.st-key-hero_search_area div[data-testid="stButton"] button { height:34px !important; min-height:34px !important; padding:3px 8px !important; margin-top:3px !important; border-radius:18px !important; border:1px solid rgba(255,255,255,.34) !important; background:rgba(255,255,255,.11) !important; color:white !important; font-size:9px !important; font-weight:500 !important; box-shadow:none !important; white-space:nowrap !important; }
.st-key-hero_search_area div[data-testid="stButton"] button:hover { background:rgba(255,255,255,.20) !important; border-color:rgba(255,255,255,.60) !important; }

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



.family-stack-title {
  color:#0a3154;
  font-size:12px;
  font-weight:800;
  margin:0 0 7px 2px;
}
.family-link {
  display:block !important;
  text-decoration:none !important;
  color:inherit !important;
  -webkit-tap-highlight-color:transparent;
}
.family-link:hover,
.family-link:focus,
.family-link:active {
  text-decoration:none !important;
  color:inherit !important;
}
.family-link .family-card {
  cursor:pointer;
  transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}
.family-link:hover .family-card {
  transform:translateX(-2px);
  border-color:#9edfd8;
  box-shadow:0 8px 24px rgba(0,150,135,.15);
}
.family-link:focus-visible .family-card {
  outline:3px solid rgba(0,191,165,.28);
  outline-offset:2px;
}
.family-card-compact {
  height:58px !important;
  min-height:58px !important;
  padding:8px 10px !important;
  margin-bottom:7px !important;
  display:flex !important;
  align-items:center !important;
  gap:9px !important;
  border-radius:12px !important;
}
.family-card-compact .family-icon {
  width:34px !important;
  height:34px !important;
  min-width:34px !important;
  border-radius:9px !important;
  margin:0 !important;
  font-size:17px !important;
}
.family-card-compact .family-card-copy {
  min-width:0;
  flex:1;
}
.family-card-compact .family-name {
  font-size:13px !important;
  line-height:16px !important;
}
.family-card-compact .family-desc {
  font-size:9px !important;
  line-height:11px !important;
  margin-top:2px !important;
  padding-right:0 !important;
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
}
.family-card-compact .family-arrow {
  position:static !important;
  flex:0 0 26px;
  width:26px !important;
  height:26px !important;
  font-size:15px !important;
}

.main-grid {
  display:grid;
  grid-template-columns:1fr;
  gap:12px;
}

.panel {
  background:rgba(255,255,255,.90);
  border:1px solid rgba(215,229,234,.96);
  border-radius:18px;
  box-shadow:0 6px 19px rgba(35,74,91,.07);
  padding:16px;
}

.ai-panel { min-height:0 !important; height:auto !important; }

.ai-panel div[data-testid="stTextInput"] {
  margin:7px 0 4px !important;
}
.ai-panel div[data-testid="stTextInput"] input {
  height:34px !important;
  min-height:34px !important;
  border-radius:9px !important;
  border:1px solid #d7e4e8 !important;
  background:#eef3f6 !important;
  color:#243f55 !important;
  font-size:10px !important;
  padding:0 10px !important;
  box-shadow:none !important;
}
.ai-panel div[data-testid="stFormSubmitButton"] button,
.ai-panel button[kind="secondaryFormSubmit"] {
  height:34px !important;
  min-height:34px !important;
  width:34px !important;
  padding:0 !important;
  border-radius:50% !important;
  border:0 !important;
  background:#08bda4 !important;
  color:white !important;
  font-size:15px !important;
  font-weight:800 !important;
}
.ai-panel div[data-testid="stFormSubmitButton"] button:hover {
  background:#009f8d !important;
  color:white !important;
}
.ai-panel div[data-testid="stButton"] button {
  min-height:26px !important;
  padding:3px 7px !important;
  border-radius:16px !important;
  font-size:8px !important;
  line-height:1.1 !important;
  white-space:nowrap !important;
}

.ai-panel div[data-testid="stButton"] button p,
.ai-panel div[data-testid="stButton"] button span {
  font-size:8px !important;
  line-height:1.1 !important;
  margin:0 !important;
}
.ai-panel .stForm {
  border:0 !important;
  padding:0 !important;
}
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

.topic-link {
  display:block;
  text-decoration:none !important;
  color:inherit !important;
  -webkit-tap-highlight-color:transparent;
  margin-bottom:8px;
}
.topic-link:hover,
.topic-link:focus,
.topic-link:active {
  text-decoration:none !important;
  color:inherit !important;
}
.topic-link .topic-card {
  cursor:pointer;
  transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}
.topic-link:hover .topic-card {
  transform:translateY(-1px);
  border-color:#9edfd8;
  box-shadow:0 6px 18px rgba(0,150,135,.12);
}
.topic-link:focus-visible .topic-card {
  outline:3px solid rgba(0,191,165,.28);
  outline-offset:2px;
}
.topic-card {
  min-height:58px;
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
  font-size:12px;
  font-weight:700;
}

.topic-count {
  color:#7b8f9d;
  font-size:9px;
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
  font-size:12px !important;
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


.small-button div.stButton > button {
  min-height:26px !important;
  font-size:8px !important;
  padding:2px 8px !important;
}

@media (max-width:1100px){ .hero-grid{grid-template-columns:170px 1fr 120px}.hero-title{font-size:35px;line-height:40px}.hero-sub{font-size:13px}.family-card{height:145px}.main-grid{grid-template-columns:1fr}.st-key-hero_search_area{max-width:760px}.st-key-hero_search_area div[data-testid="stButton"] button{font-size:8px !important} }
@media (max-width:800px){ .block-container{padding:0 10px 14px !important}.hero{height:340px;min-height:340px;margin:0 -10px;padding:14px 16px 12px}.hero-grid{grid-template-columns:1fr;height:auto;text-align:center}.brand{text-align:center}.hpe-logo{margin-left:auto;margin-right:auto}.brand-rule{margin-left:auto;margin-right:auto}.brand-copy{font-size:9px}.hero-center{margin-top:14px}.hero-title{font-size:29px;line-height:33px;letter-spacing:-1px;padding:0 8px}.hero-sub{font-size:11px;line-height:15px;padding:0 15px;margin-top:5px}.hero-right{display:none}.st-key-hero_search_area{width:calc(100% - 20px) !important;max-width:none !important;margin:-110px auto 10px !important}.st-key-hero_search_area div[data-testid="stTextInput"] input{height:50px !important;font-size:12px !important;padding:0 13px !important}.st-key-hero_search_area div[data-testid="stFormSubmitButton"] button{width:48px !important;height:48px !important;min-height:48px !important;font-size:22px !important}.st-key-hero_search_area .try-label{display:block;text-align:center;margin-top:7px;font-size:9px}.st-key-hero_search_area div[data-testid="stHorizontalBlock"]{gap:4px !important}.st-key-hero_search_area div[data-testid="stButton"] button{height:auto !important;min-height:31px !important;font-size:7px !important;padding:4px !important;white-space:normal !important;line-height:9px !important}.family-card{height:142px}.doc-grid{grid-template-columns:repeat(2,1fr)}.topic-grid{grid-template-columns:1fr}.main-grid{display:block}.panel{padding:12px;border-radius:14px}.bottom-strip{grid-template-columns:repeat(2,1fr);height:auto;padding:7px 0}.stat{min-height:48px;border-right:0}.stat:last-child{grid-column:1 / -1} }
@media (max-width:480px){ .hero{height:360px;min-height:360px}.hero-title{font-size:25px;line-height:29px}.hero-sub{font-size:10px}.st-key-hero_search_area{margin-top:-116px !important}.family-card{height:136px;padding:11px 10px}.family-name{font-size:14px}.family-desc{font-size:9px;line-height:12px}.family-icon{width:34px;height:34px;font-size:18px}.doc-grid{grid-template-columns:1fr}.topic-grid{grid-template-columns:1fr} }

/* Smaller answer-choice text — keeps long suggested questions compact. */
.ai-panel div[data-testid="stButton"] button,
.ai-panel div[data-testid="stButton"] button p,
.ai-panel div[data-testid="stButton"] button span {
  font-size:8px !important;
  line-height:1.1 !important;
}

@media (max-width:800px) {
  .ai-panel div[data-testid="stButton"] button,
  .ai-panel div[data-testid="stButton"] button p,
  .ai-panel div[data-testid="stButton"] button span {
    font-size:7px !important;
    line-height:1.05 !important;
  }
}

/* The AI Assistant uses form submission only.
   No standalone "Ask this question" button is rendered. */


/* Admin gear: visually anchored INSIDE the dark-teal hero, directly above
   the “Accelerating what's next together” message. The Streamlit popover
   remains the real accessible/password-protected control. */
.st-key-hero_admin_gear {
  position:absolute !important;
  top:18px !important;
  right:18px !important;
  width:48px !important;
  min-width:48px !important;
  max-width:48px !important;
  height:48px !important;
  margin:0 !important;
  padding:0 !important;
  z-index:1000 !important;
}
.st-key-hero_admin_gear > div,
.st-key-hero_admin_gear [data-testid="stVerticalBlock"] {
  width:48px !important;
  min-width:48px !important;
  padding:0 !important;
  margin:0 !important;
}
.st-key-hero_admin_gear div[data-testid="stPopover"] {
  width:48px !important;
}
.st-key-hero_admin_gear div[data-testid="stPopover"] > button {
  width:46px !important;
  height:46px !important;
  min-width:46px !important;
  min-height:46px !important;
  padding:0 !important;
  margin:0 !important;
  border-radius:50% !important;
  background:#00bfa5 !important;
  border:2px solid rgba(0,238,216,.72) !important;
  color:#ffffff !important;
  box-shadow:0 4px 14px rgba(0,0,0,.25) !important;
  font-size:25px !important;
  line-height:1 !important;
  cursor:pointer !important;
}
.st-key-hero_admin_gear div[data-testid="stPopover"] > button:hover,
.st-key-hero_admin_gear div[data-testid="stPopover"] > button:focus-visible {
  background:#00d7bd !important;
  border-color:#55f2df !important;
  color:#ffffff !important;
  outline:3px solid rgba(0,230,210,.28) !important;
  outline-offset:3px !important;
}
.st-key-hero_admin_gear [data-testid="stPopoverBody"] {
  right:0 !important;
  left:auto !important;
  top:54px !important;
  z-index:2000 !important;
}
.ai-panel {
  min-height:0 !important;
  height:auto !important;
}
.ai-panel .ai-head {
  margin-bottom:6px !important;
}


.chatbot-prompt {
  margin:9px 0 6px 12px;
  color:#526c7d;
  font-size:9px;
  font-weight:600;
}
.chatbot-related-label {
  margin:8px 0 5px 12px;
  color:#173a56;
  font-size:9px;
  font-weight:800;
}
.exact-answer-label {
  display:inline-block;
  margin-bottom:7px;
  color:#008f7b;
  font-size:10px;
  font-weight:800;
  letter-spacing:.6px;
}


@media (max-width:800px){
  .st-key-hero_admin_gear {
    top:13px !important;
    right:14px !important;
  }
  .family-card-compact {
    height:62px !important;
    min-height:62px !important;
  }
  .family-card-compact .family-desc {
    white-space:normal !important;
    display:-webkit-box;
    -webkit-line-clamp:2;
    -webkit-box-orient:vertical;
  }
}



/* Functional family tiles: compact tiles with LARGE icons, not large tiles. */
.st-key-family_tile_0,
.st-key-family_tile_1,
.st-key-family_tile_2,
.st-key-family_tile_3,
.st-key-family_tile_4,
.st-key-family_tile_5 {
  position:relative !important;
  margin:0 !important;
  padding:0 !important;
}
.st-key-family_tile_0 > div,
.st-key-family_tile_1 > div,
.st-key-family_tile_2 > div,
.st-key-family_tile_3 > div,
.st-key-family_tile_4 > div,
.st-key-family_tile_5 > div { position:relative !important; }
.family-tile-visual {
  position:relative;
  height:60px;
  min-height:60px;
  width:100%;
  border:1px solid #d5e4ea;
  border-radius:12px;
  background:rgba(255,255,255,.95);
  box-shadow:0 4px 14px rgba(36,76,92,.06);
  display:flex;
  align-items:center;
  gap:12px;
  padding:7px 42px 7px 12px;
  pointer-events:none;
}
.family-tile-visual .family-icon {
  width:38px !important;
  height:38px !important;
  min-width:38px !important;
  border-radius:10px !important;
  margin:0 !important;
  font-size:24px !important;
  line-height:1 !important;
  display:flex !important;
  align-items:center !important;
  justify-content:center !important;
}
.family-tile-copy { min-width:0; flex:1; }
.family-tile-visual .family-name {
  font-size:13px !important;
  line-height:16px !important;
  font-weight:800 !important;
  color:#0a3154 !important;
}
.family-tile-visual .family-desc {
  margin-top:2px !important;
  padding:0 !important;
  font-size:8px !important;
  line-height:10px !important;
  color:#536f82 !important;
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
}
.family-tile-visual .family-arrow {
  position:absolute !important;
  right:11px !important;
  top:17px !important;
  width:27px !important;
  height:27px !important;
  border-radius:50% !important;
  display:flex !important;
  align-items:center !important;
  justify-content:center !important;
  font-size:16px !important;
  font-weight:800 !important;
}
[class*="st-key-family_tile_btn_"] {
  position:absolute !important;
  inset:0 !important;
  z-index:5 !important;
  margin:0 !important;
  padding:0 !important;
}
[class*="st-key-family_tile_btn_"] button {
  position:absolute !important;
  inset:0 !important;
  width:100% !important;
  height:60px !important;
  min-height:60px !important;
  margin:0 !important;
  padding:0 !important;
  border-radius:12px !important;
  border:2px solid transparent !important;
  background:transparent !important;
  color:transparent !important;
  box-shadow:none !important;
  font-size:0 !important;
  line-height:0 !important;
}
[class*="st-key-family_tile_btn_"] button:hover {
  background:rgba(0,191,165,.025) !important;
  border-color:#8ed8cf !important;
  box-shadow:0 7px 20px rgba(0,150,135,.10) !important;
  transform:none !important;
}
[class*="st-key-family_tile_btn_"] button:focus-visible {
  outline:3px solid rgba(0,191,165,.35) !important;
  outline-offset:2px !important;
}
@media (max-width:800px) {
  .family-tile-visual { height:60px; min-height:60px; padding:7px 39px 7px 10px; gap:10px; }
  .family-tile-visual .family-icon { width:38px !important; height:38px !important; min-width:38px !important; font-size:24px !important; }
  .family-tile-visual .family-name { font-size:12px !important; line-height:14px !important; }
  .family-tile-visual .family-desc { font-size:7px !important; line-height:9px !important; }
  [class*="st-key-family_tile_btn_"] button { height:60px !important; min-height:60px !important; }
}

.st-key-home_ai_exact_answer_box,
[class*="st-key-answer_"][class*="_exact_answer_box"] {
  margin:9px 0 8px 12px !important;
  padding:0 14px 12px !important;
  border:1px solid #cfe6e1 !important;
  border-left:4px solid #00bfa5 !important;
  border-radius:0 12px 12px 12px !important;
  background:#f7fcfb !important;
  box-shadow:0 2px 9px rgba(0,130,115,.04) !important;
}
.exact-section-title {
  margin:12px 0 5px;
  color:#008f7b;
  font-size:8px;
  font-weight:800;
  letter-spacing:.65px;
}
/* More readable exact answer presentation. */
.exact-answer {
  margin:9px 0 8px 12px;
  padding:14px 16px 15px;
  border:1px solid #cfe6e1;
  border-left:4px solid #00bfa5;
  border-radius:0 12px 12px 12px;
  background:#f7fcfb;
}
.exact-answer-header {
  display:flex;
  align-items:center;
  gap:10px;
  padding-bottom:10px;
  margin-bottom:10px;
  border-bottom:1px solid #dfecea;
}
.exact-answer-icon {
  width:28px;
  height:28px;
  border-radius:50%;
  background:#08bda4;
  color:#fff;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:16px;
  font-weight:800;
}
.exact-answer-label {
  margin:0 !important;
  color:#008f7b !important;
  font-size:9px !important;
  font-weight:800 !important;
  letter-spacing:.7px !important;
}
.exact-answer-title {
  margin-top:2px;
  color:#173a56;
  font-size:10px;
  font-weight:700;
}
.exact-answer-body {
  color:#243f55;
  font-size:13px;
  line-height:1.65;
  white-space:normal;
}
.summary-paragraph {
  color:#243f55;
  font-size:12px;
  line-height:1.7;
  margin:0 0 8px;
  padding:0;
}
.summary-subtitle {
  color:#173a56;
  font-size:11px;
  font-weight:800;
  margin:8px 0 4px;
}
.summary-bullet {
  display:flex;
  gap:8px;
  color:#324e63;
  font-size:11px;
  line-height:1.5;
  margin:4px 0;
}
.summary-bullet > span {
  color:#00a991;
  font-weight:900;
  flex:0 0 auto;
}
.related-empty {
  color:#7a8d99;
  font-size:9px;
  padding:8px 0 2px;
}
.msg-label {
  display:inline-block;
  margin-right:8px;
  color:#7890a0;
  font-size:8px;
  font-weight:800;
  letter-spacing:.5px;
}

@media (max-width:800px) {
  .st-key-hero_admin_gear {
    top:12px !important;
    right:12px !important;
    width:48px !important;
    min-width:48px !important;
  }
  .st-key-hero_admin_gear div[data-testid="stPopover"] > button {
    width:44px !important;
    height:44px !important;
    min-width:44px !important;
    min-height:44px !important;
    font-size:23px !important;
  }
  .exact-answer { margin-left:0; }
  .exact-answer-body { font-size:11px; line-height:1.6; }
}

/* ============================================================
   EXACT ANSWER / AI RESPONSE — READABLE REFERENCE SCALE
   Keep the typography visually consistent across the entire
   answer section shown in the supplied reference image.
   ============================================================ */
.st-key-home_ai_exact_answer_box .user-msg,
[class*="st-key-answer_"][class*="_exact_answer_box"] .user-msg {
  font-size:10px !important;
  line-height:1.5 !important;
}
.st-key-home_ai_exact_answer_box .msg-label,
[class*="st-key-answer_"][class*="_exact_answer_box"] .msg-label {
  font-size:8px !important;
}
.st-key-home_ai_exact_answer_box .exact-answer-label,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-answer-label {
  font-size:9px !important;
}
.st-key-home_ai_exact_answer_box .exact-answer-title,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-answer-title {
  font-size:10px !important;
  line-height:1.35 !important;
}
.st-key-home_ai_exact_answer_box .exact-answer-body,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-answer-body {
  font-size:11px !important;
  line-height:1.6 !important;
}
.st-key-home_ai_exact_answer_box .exact-section-title,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-section-title {
  font-size:9px !important;
  line-height:1.4 !important;
}
.st-key-home_ai_exact_answer_box .summary-paragraph,
[class*="st-key-answer_"][class*="_exact_answer_box"] .summary-paragraph {
  font-size:10px !important;
  line-height:1.6 !important;
}
.st-key-home_ai_exact_answer_box .summary-subtitle,
[class*="st-key-answer_"][class*="_exact_answer_box"] .summary-subtitle {
  font-size:10px !important;
}
.st-key-home_ai_exact_answer_box .summary-bullet,
[class*="st-key-answer_"][class*="_exact_answer_box"] .summary-bullet {
  font-size:10px !important;
  line-height:1.5 !important;
}
.st-key-home_ai_exact_answer_box .ai-step,
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-step {
  font-size:10px !important;
}
.st-key-home_ai_exact_answer_box .ai-step > div:last-child,
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-step > div:last-child {
  font-size:10px !important;
  line-height:1.5 !important;
}
.st-key-home_ai_exact_answer_box .ai-num,
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-num {
  font-size:9px !important;
}
.st-key-home_ai_exact_answer_box .ai-doc,
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-doc {
  font-size:10px !important;
}
.st-key-home_ai_exact_answer_box .ai-doc div[style*="font-size:9px"],
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-doc div[style*="font-size:9px"] {
  font-size:10px !important;
}
.st-key-home_ai_exact_answer_box .ai-doc div[style*="font-size:8px"],
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-doc div[style*="font-size:8px"] {
  font-size:9px !important;
}

/* Keep the three action controls readable without making them oversized. */
.st-key-home_ai_exact_answer_box ~ div div[data-testid="stButton"] button,
[class*="st-key-answer_"][class*="_exact_answer_box"] ~ div div[data-testid="stButton"] button {
  font-size:10px !important;
  min-height:30px !important;
}

/* FINAL AI ANSWER TYPOGRAPHY OVERRIDE */
.st-key-home_ai_exact_answer_box .user-msg,
[class*="st-key-answer_"][class*="_exact_answer_box"] .user-msg {
  font-size:11px !important;
  line-height:1.55 !important;
}
.st-key-home_ai_exact_answer_box .msg-label,
[class*="st-key-answer_"][class*="_exact_answer_box"] .msg-label {
  font-size:8px !important;
  line-height:1.2 !important;
}
.st-key-home_ai_exact_answer_box .exact-answer-label,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-answer-label {
  font-size:9px !important;
  line-height:1.2 !important;
}
.st-key-home_ai_exact_answer_box .exact-answer-title,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-answer-title {
  font-size:11px !important;
  line-height:1.35 !important;
}
.st-key-home_ai_exact_answer_box .exact-answer-body,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-answer-body {
  font-size:14px !important;
  line-height:1.65 !important;
}
.st-key-home_ai_exact_answer_box .exact-section-title,
[class*="st-key-answer_"][class*="_exact_answer_box"] .exact-section-title {
  font-size:10px !important;
  line-height:1.35 !important;
}
.st-key-home_ai_exact_answer_box .summary-paragraph,
[class*="st-key-answer_"][class*="_exact_answer_box"] .summary-paragraph {
  font-size:13px !important;
  line-height:1.65 !important;
}
.st-key-home_ai_exact_answer_box .summary-subtitle,
[class*="st-key-answer_"][class*="_exact_answer_box"] .summary-subtitle {
  font-size:11px !important;
}
.st-key-home_ai_exact_answer_box .summary-bullet,
[class*="st-key-answer_"][class*="_exact_answer_box"] .summary-bullet {
  font-size:12px !important;
  line-height:1.55 !important;
}
.st-key-home_ai_exact_answer_box .ai-step > div:last-child,
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-step > div:last-child {
  font-size:12px !important;
  line-height:1.55 !important;
}
.st-key-home_ai_exact_answer_box .ai-num,
[class*="st-key-answer_"][class*="_exact_answer_box"] .ai-num {
  font-size:10px !important;
}
@media (max-width:800px) {
  .st-key-home_ai_exact_answer_box .exact-answer-body,
  [class*="st-key-answer_"][class*="_exact_answer_box"] .exact-answer-body {
    font-size:13px !important;
    line-height:1.6 !important;
  }
  .st-key-home_ai_exact_answer_box .summary-paragraph,
  [class*="st-key-answer_"][class*="_exact_answer_box"] .summary-paragraph {
    font-size:12px !important;
  }
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
        </div>
        <div class="hero-right">
          Accelerating<br>what's next<br><span class="next">together</span>
          <div class="dash">—</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    with st.container(key="hero_admin_gear"):
        admin_login()

    with st.container(key="hero_search_area"):
        with st.form("hero_search_form", clear_on_submit=False):
            c1, c2 = st.columns([0.93, 0.07], gap="small", vertical_alignment="center")
            with c1:
                pending_global_query = st.session_state.pop("global_query_pending", None)
                if pending_global_query is not None:
                    st.session_state["global_query"] = pending_global_query
                st.text_input(
                    "Global search",
                    placeholder="Ask a question or search for a document...",
                    key="global_query",
                    label_visibility="collapsed"
                )
            with c2:
                submitted = st.form_submit_button("→", use_container_width=True)
        if submitted:
            value = st.session_state.get("global_query", "").strip()
            if value:
                st.session_state["home_ai_query"] = value
                st.session_state["home_ai_submitted"] = value
                st.session_state["home_ai_action"] = None
            st.session_state.view = "home"
            st.rerun()



def open_group(group):
    st.session_state.view = "group"
    st.session_state.selected_group = group
    st.session_state.selected_topic = None
    st.rerun()


def render_family_cards():
    """Render compact, fully clickable family tiles without a second button row."""
    from urllib.parse import quote

    st.markdown(
        '<div class="family-stack-title">HPE & Aruba Product Families</div>',
        unsafe_allow_html=True
    )

    for group, data in PRODUCT_GROUPS.items():
        # Use a normal in-app URL instead of a Streamlit button. This keeps the
        # clickable hit area exactly on top of the visual tile and prevents the
        # button from rendering as a separate white row underneath it.
        family_param = quote(group, safe="")
        st.markdown(
            f"""
            <a class="family-link" href="?family={family_param}"
               aria-label="Open {eh(group)} product family">
              <div class="family-card family-card-compact">
                <div class="family-icon {data['class']}">{data['icon']}</div>
                <div class="family-card-copy">
                  <div class="family-name">{eh(group)}</div>
                  <div class="family-desc">{eh(data['description'])}</div>
                </div>
                <div class="family-arrow arrow-{data['class']}">→</div>
              </div>
            </a>
            """,
            unsafe_allow_html=True
        )

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
    """Compatibility wrapper that renders the real functional AI Assistant."""
    safe = re.sub(r"[^A-Za-z0-9]+", "_", (family or "global")).strip("_").lower()
    prefix = f"answer_{safe or 'global'}"
    st.session_state[f"{prefix}_query"] = query
    st.session_state[f"{prefix}_submitted"] = query
    st.session_state[f"{prefix}_action"] = None
    render_ai_assistant(
        default_query=query,
        family=family,
        key_prefix=prefix,
    )


def render_ai_assistant(default_query="",
                         family=None, key_prefix="home_ai"):
    """Option-based HPE AI chatbot with exact-answer retrieval."""
    q_key = f"{key_prefix}_query"
    submitted_key = f"{key_prefix}_submitted"
    selected_key = f"{key_prefix}_selected"
    action_key = f"{key_prefix}_action"

    # The assistant starts blank on the home page. A non-empty default_query
    # is only used by an explicit programmatic call such as render_ai_answer().
    if q_key not in st.session_state:
        st.session_state[q_key] = default_query or ""

    if submitted_key not in st.session_state:
        st.session_state[submitted_key] = default_query or ""

    # Never resurrect the old demo question if it exists in a stale session.
    previous_demo = "How do I troubleshoot ClearPass licensing issues?"
    if not default_query:
        if st.session_state.get(q_key) == previous_demo:
            st.session_state[q_key] = ""
        if st.session_state.get(submitted_key) == previous_demo:
            st.session_state[submitted_key] = ""

    st.session_state.setdefault(selected_key, None)
    st.session_state.setdefault(action_key, None)

    st.markdown("""
    <div class="panel ai-panel">
      <div class="ai-head">
        <div class="robot">🤖</div>
        <div>
          <div class="panel-title">HPE AI Assistant <span class="beta">BETA</span></div>
          <div class="panel-sub">Ask a question. Choose the closest result. Get the exact answer.</div>
        </div>
        <div class="powered">✦ Powered by HPE Knowledge</div>
      </div>
    """, unsafe_allow_html=True)

    with st.form(f"{key_prefix}_question_form", clear_on_submit=False):
        c1, c2 = st.columns([0.94, 0.06], gap="small", vertical_alignment="center")
        with c1:
            st.text_input(
                "Ask HPE AI",
                key=q_key,
                placeholder="Ask a question...",
                label_visibility="collapsed",
                help="Type your question and press Enter, or select the arrow to submit."
            )
        with c2:
            submit = st.form_submit_button("→", use_container_width=True)

    if submit:
        value = st.session_state[q_key].strip()
        if value:
            st.session_state[submitted_key] = value
            st.session_state[selected_key] = None
            st.session_state[action_key] = None
            st.session_state.global_query_pending = value
            st.rerun()

    query = st.session_state[submitted_key].strip()

    # Keep the assistant blank until the user enters a question.
    if not query:
        st.markdown("</div>", unsafe_allow_html=True)
        return

    records, docs = search(query, family=family, limit=6)

    if not records and docs:
        d = docs[0]
        excerpt = d["content"][:1200].replace("\n", " ")
        records = [{
            "kb_id": f"DOC-{d['id']}",
            "family": d["family"],
            "topic": d["topic"],
            "question": query,
            "answer": excerpt + ("…" if len(d["content"]) > 1200 else ""),
            "steps": "",
            "keywords": query,
            "source": d["title"],
        }]

    if not records:
        st.warning(
            "I couldn't find a confident answer. Try adding the product name, "
            "feature, model, version, or exact error message."
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    options = records[:3]
    selected_id = st.session_state.get(selected_key)
    selected = next((r for r in options if r["kb_id"] == selected_id), None)

    if len(options) > 1 and selected is None:
        st.markdown(
            '<div class="chatbot-prompt">I found these possible answers. Choose the one that best matches your question:</div>',
            unsafe_allow_html=True
        )
        for i, r in enumerate(options):
            if st.button(
                r["question"],
                key=f"{key_prefix}_option_{i}_{r['kb_id']}",
                use_container_width=True
            ):
                st.session_state[selected_key] = r["kb_id"]
                st.session_state[action_key] = None
                st.rerun()
        selected = options[0]
    else:
        selected = selected or options[0]

    # Keep the complete answer experience inside one bordered answer box.
    with st.container(key=f"{key_prefix}_exact_answer_box"):
        st.markdown(
            f"""
            <div class="user-msg"><span class="msg-label">QUESTION</span>{eh(query)}</div>
            <div class="exact-answer-header">
              <span class="exact-answer-icon">✓</span>
              <div>
                <div class="exact-answer-label">EXACT ANSWER</div>
                <div class="exact-answer-title">{eh(selected["topic"])}</div>
              </div>
            </div>
            <div class="exact-answer-body">{eh(selected["answer"])}</div>
            """,
            unsafe_allow_html=True
        )

        a1, a2, a3 = st.columns(3, gap="small")
        with a1:
            if st.button("✦ Summarize", key=f"{key_prefix}_summary", use_container_width=True):
                st.session_state[action_key] = "summary"
                st.rerun()
        with a2:
            if st.button("⌕ Troubleshooting steps", key=f"{key_prefix}_steps", use_container_width=True):
                st.session_state[action_key] = "steps"
                st.rerun()
        with a3:
            if st.button("▤ Related knowledge", key=f"{key_prefix}_related", use_container_width=True):
                st.session_state[action_key] = "related"
                st.rerun()

        action = st.session_state[action_key]

        # The answer box defaults to the direct answer only. The three actions
        # deliberately reveal their content inside the same answer container.
        if action == "summary":
            answer_text = (selected.get("answer") or "").strip()
            step_lines = []
            for line in (selected.get("steps") or "").splitlines():
                clean = re.sub(r"^\s*\d+\.\s*", "", line).strip()
                if clean:
                    step_lines.append(clean)

            # Keep the summary readable: one concise paragraph followed by
            # only the most useful action points when the source contains steps.
            summary_paragraph = re.sub(r"\s+", " ", answer_text).strip()
            st.markdown('<div class="exact-section-title">SUMMARY</div>', unsafe_allow_html=True)
            st.markdown(
                f'<div class="summary-paragraph">{eh(summary_paragraph)}</div>',
                unsafe_allow_html=True
            )
            if step_lines:
                st.markdown('<div class="summary-subtitle">Key points</div>', unsafe_allow_html=True)
                for item in step_lines[:5]:
                    st.markdown(
                        f'<div class="summary-bullet"><span>•</span><div>{eh(item)}</div></div>',
                        unsafe_allow_html=True
                    )

        elif action == "steps":
            st.markdown('<div class="exact-section-title">TROUBLESHOOTING STEPS</div>', unsafe_allow_html=True)
            for i, line in enumerate((selected.get("steps") or "").splitlines(), 1):
                clean = re.sub(r"^\s*\d+\.\s*", "", line)
                if clean.strip():
                    st.markdown(
                        f'<div class="ai-step"><div class="ai-num">{i}</div>'
                        f'<div style="font-size:9px;color:#324e63;padding-top:3px;line-height:1.45;">{eh(clean)}</div></div>',
                        unsafe_allow_html=True
                    )

        elif action == "related":
            st.markdown('<div class="exact-section-title">RELATED KNOWLEDGE</div>', unsafe_allow_html=True)

            # Related KB records
            for r in records[1:5]:
                st.markdown(
                    f'<div class="ai-doc"><div class="pdf-icon">▤</div><div style="flex:1;">'
                    f'<div style="font-size:9px;font-weight:800;color:#173a56;">{eh(r["question"])}</div>'
                    f'<div style="font-size:8px;color:#8a9aa5;">{eh(r["family"])} • {eh(r["topic"])} • {eh(r["kb_id"])}</div>'
                    f'</div></div>',
                    unsafe_allow_html=True
                )

            # PDFs belong ONLY to Related knowledge; they are no longer
            # rendered underneath the default exact answer.
            if docs:
                for d in docs[:4]:
                    st.markdown(
                        f'<div class="ai-doc"><div class="pdf-icon">▤</div><div style="flex:1;">'
                        f'<div style="font-size:9px;font-weight:800;color:#173a56;">{eh(d["title"])}</div>'
                        f'<div style="font-size:8px;color:#8a9aa5;">⌁ PDF • Indexed • {eh(d["family"])}</div>'
                        f'</div></div>',
                        unsafe_allow_html=True
                    )
                    if st.button("↗ View source document", key=f"{key_prefix}_view_source_{d['id']}", use_container_width=True):
                        st.session_state.selected_document = d["id"]
                        st.session_state.view = "document"
                        st.rerun()
            elif len(records) <= 1:
                st.markdown(
                    '<div class="related-empty">No additional PDF source was indexed for this answer yet.</div>',
                    unsafe_allow_html=True
                )

    # No follow-up input: the original question field remains the single active question entry.

    st.markdown("</div>", unsafe_allow_html=True)



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
                    st.session_state.global_query_pending = title
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
    # Sync a NEW global-search query into the AI assistant, but do not
    # overwrite the selected action on every Streamlit rerun.  Previously,
    # clicking Summarize / Troubleshooting steps / Related knowledge set
    # home_ai_action and then render_home immediately reset it to None,
    # making the buttons appear non-functional.
    current_ai_query = st.session_state.get("home_ai_submitted", "").strip()
    if query and query != current_ai_query:
        st.session_state["home_ai_query"] = query
        st.session_state["home_ai_submitted"] = query
        st.session_state["home_ai_selected"] = None
        st.session_state["home_ai_action"] = None

    left, right = st.columns([1.72, 0.78], gap="medium")

    with left:
        render_ai_assistant(
            default_query=query,
            key_prefix="home_ai",
        )

    with right:
        render_family_cards()

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
        try:
            st.query_params.clear()
        except Exception:
            pass
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

    # Product directory cards were intentionally removed from this page.
    # Topics are the primary drill-down navigation and each tile is clickable.
    st.markdown("""
    <div class="panel" style="margin-bottom:10px;">
      <div class="panel-title">Topics & Solutions</div>
      <div class="panel-sub">Select a topic to view its related AI answers and sources.</div>
      <div style="height:8px"></div>
    """, unsafe_allow_html=True)

    from urllib.parse import quote
    topic_cols = st.columns(3, gap="small")
    family_param = quote(group, safe="")
    for i, topic in enumerate(data["topics"]):
        with topic_cols[i % 3]:
            topic_param = quote(topic, safe="")
            st.markdown(
                f"""
                <a class="topic-link"
                   href="?family={family_param}&topic={topic_param}"
                   aria-label="Open {eh(topic)}">
                  <div class="topic-card">
                    <div class="topic-icon">{eh(data['icon'])}</div>
                    <div style="flex:1;min-width:0">
                      <div class="topic-name">{eh(topic)}</div>
                      <div class="topic-count">View related knowledge</div>
                    </div>
                    <div style="color:#6b8191;font-size:18px;line-height:1;">›</div>
                  </div>
                </a>
                """,
                unsafe_allow_html=True
            )

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
                  <a href="{eh(url)}" target="_self">Open source →</a>
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
# ADMIN ACCESS + SOP AUTHORING
# ============================================================
def get_admin_password():
    """Read the admin password only from Streamlit secrets."""
    try:
        if "ADMIN_PASSWORD" in st.secrets:
            return str(st.secrets["ADMIN_PASSWORD"])
        if "ADMIN_PIN" in st.secrets:
            return str(st.secrets["ADMIN_PIN"])
    except Exception:
        pass
    return ""


def admin_login():
    """Password gate for the admin area. Password never lives in source code."""
    with st.popover("⚙", help="Knowledge Base Admin"):
        st.markdown("#### Admin")
        st.caption("Enter the admin password configured in Streamlit Secrets.")

        with st.form("admin_login_form"):
            password = st.text_input(
                "Password",
                type="password",
                autocomplete="current-password",
                label_visibility="collapsed",
                placeholder="Admin password"
            )
            login = st.form_submit_button("Sign in", use_container_width=True)

        if login:
            expected = get_admin_password()
            if expected and password == expected:
                st.session_state.admin_authenticated = True
                st.session_state.view = "admin"
                st.rerun()
            elif not expected:
                st.error("ADMIN_PASSWORD is not configured in Streamlit Secrets.")
            else:
                st.error("Incorrect admin password.")


def create_ai_ready_sop(
    family, product, topic, question, direct_answer,
    prerequisites, steps, verification, escalation, keywords,
    source_title="", source_url=""
):
    """
    Store one atomic SOP record designed for deterministic AI retrieval.
    The record is intentionally self-contained:
    Question -> Direct Answer -> Steps -> Verification -> Escalation.
    """
    now = datetime.now().isoformat(timespec="seconds")
    kb_id = f"SOP-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"

    # Build a normalized searchable text field in the existing answer/steps
    # columns without changing the database schema.
    normalized_steps = steps.strip()
    if prerequisites.strip():
        normalized_steps = (
            "Prerequisites:\n" + prerequisites.strip()
            + "\n\nProcedure:\n" + normalized_steps
        )
    if verification.strip():
        normalized_steps += "\n\nVerification:\n" + verification.strip()
    if escalation.strip():
        normalized_steps += "\n\nEscalation:\n" + escalation.strip()

    source = source_title.strip() or "Admin-created SOP"
    if source_url.strip():
        source += f" | {source_url.strip()}"

    conn = db()
    conn.execute("""
        INSERT INTO kb_records
        (kb_id, family, topic, question, answer, steps, keywords, source, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        kb_id,
        family,
        f"{product} — {topic}",
        question.strip(),
        direct_answer.strip(),
        normalized_steps,
        ", ".join([x.strip() for x in keywords.split(",") if x.strip()]),
        source,
        now
    ))
    conn.commit()
    conn.close()
    search.clear()
    return kb_id


def render_admin():
    if st.button("← Back to Knowledge Base", key="admin_back"):
        st.session_state.view = "home"
        st.rerun()

    st.markdown("""
    <div class="family-banner">
      <h1>⚙ Knowledge Base Admin</h1>
      <p>Create AI-ready SOPs or upload PDF source documents. Every SOP is stored as an atomic question-and-answer record for accurate retrieval.</p>
    </div>
    """, unsafe_allow_html=True)

    create_tab, upload_tab = st.tabs(["Create AI-Ready SOP", "Upload PDF"])

    with create_tab:
        st.markdown("""
        <div class="panel" style="margin-bottom:12px;">
          <div class="panel-title">AI Extraction Format</div>
          <div class="panel-sub">
            Use one question per SOP. Put the exact answer first, then prerequisites,
            numbered steps, verification, escalation criteria, and search terms.
            This structure keeps retrieval precise and prevents unrelated procedures
            from being mixed into one answer.
          </div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("create_sop_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                family = st.selectbox("Product Family", list(PRODUCT_GROUPS.keys()))
                product = st.text_input("Product / Platform",
                                        placeholder="e.g. Aruba ClearPass Policy Manager")
                topic = st.text_input("Topic / Feature",
                                      placeholder="e.g. RADIUS Authentication Failure")
                question = st.text_input(
                    "Exact User Question",
                    placeholder="e.g. How do I troubleshoot a ClearPass RADIUS authentication failure?"
                )
            with c2:
                source_title = st.text_input(
                    "Source Title (optional)",
                    placeholder="e.g. ClearPass RADIUS Troubleshooting SOP"
                )
                source_url = st.text_input(
                    "Source URL (optional)",
                    placeholder="https://..."
                )
                keywords = st.text_input(
                    "Search Keywords",
                    placeholder="ClearPass, RADIUS, authentication, timeout, Access Tracker"
                )

            direct_answer = st.text_area(
                "Direct Answer — write the exact answer the AI should return",
                height=130,
                placeholder=(
                    "State the answer directly and completely. Include scope or "
                    "version limitations when they matter. Avoid introductions."
                )
            )

            prerequisites = st.text_area(
                "Prerequisites / Required Information",
                height=90,
                placeholder=(
                    "Example:\n"
                    "• Identify the ClearPass server.\n"
                    "• Record the request timestamp.\n"
                    "• Have access to Access Tracker."
                )
            )

            steps = st.text_area(
                "Procedure — one step per line",
                height=170,
                placeholder=(
                    "1. Open Access Tracker.\n"
                    "2. Locate the affected authentication request.\n"
                    "3. Confirm the selected service.\n"
                    "4. Review the authentication failure reason.\n"
                    "5. Correct the identified configuration issue.\n"
                    "6. Retest the authentication request."
                )
            )

            verification = st.text_area(
                "Verification / Expected Result",
                height=100,
                placeholder=(
                    "Describe exactly how the agent confirms the issue is resolved."
                )
            )

            escalation = st.text_area(
                "Escalation Criteria",
                height=100,
                placeholder=(
                    "State when the case must be escalated and exactly what evidence "
                    "must accompany the escalation."
                )
            )

            submitted = st.form_submit_button(
                "Create SOP & Add to AI Knowledge",
                type="primary",
                use_container_width=True
            )

        if submitted:
            required = {
                "Product / Platform": product,
                "Topic / Feature": topic,
                "Exact User Question": question,
                "Direct Answer": direct_answer,
                "Procedure": steps,
            }
            missing = [k for k, v in required.items() if not v.strip()]
            if missing:
                st.error("Complete these required fields: " + ", ".join(missing))
            else:
                kb_id = create_ai_ready_sop(
                    family, product, topic, question, direct_answer,
                    prerequisites, steps, verification, escalation,
                    keywords, source_title, source_url
                )
                st.success(f"SOP created successfully: {kb_id}")
                st.info("The new SOP is immediately searchable by the AI Assistant.")

    with upload_tab:
        c1, c2 = st.columns(2)
        with c1:
            upload_family = st.selectbox(
                "Product Family",
                list(PRODUCT_GROUPS.keys()),
                key="admin_upload_family"
            )
        with c2:
            upload_topic = st.text_input(
                "Topic",
                placeholder="e.g. ClearPass Licensing",
                key="admin_upload_topic"
            )

        file = st.file_uploader(
            "Upload PDF source document",
            type=["pdf"],
            help="PDF text is extracted and indexed into the Knowledge Base."
        )

        if file and st.button("Upload & Index PDF", type="primary", use_container_width=True):
            ok, msg = index_pdf(file, upload_family, upload_topic or "General")
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    st.markdown("### Current AI-ready SOPs")
    conn = db()
    sop_rows = conn.execute("""
        SELECT kb_id, family, topic, question, source, created_at
        FROM kb_records
        ORDER BY id DESC
    """).fetchall()
    conn.close()

    for r in sop_rows:
        with st.expander(f"{r['kb_id']} — {r['question']}"):
            st.write(f"**Family:** {r['family']}")
            st.write(f"**Topic:** {r['topic']}")
            st.write(f"**Source:** {r['source']}")
            st.caption(r["created_at"])



# ============================================================
# ROUTER
# ============================================================
for key, default in [
    ("view","home"),
    ("selected_group",None),
    ("selected_topic",None),
    ("selected_document",None),
    ("global_query",""),
    ("admin_authenticated",False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# Query parameters are used for tile navigation and as a fallback admin route.
try:
    family_param = st.query_params.get("family")
    topic_param = st.query_params.get("topic")
    if family_param:
        from urllib.parse import unquote_plus
        decoded_family = unquote_plus(str(family_param))
        decoded_topic = unquote_plus(str(topic_param)) if topic_param else None
        if decoded_family in PRODUCT_GROUPS:
            st.session_state.selected_group = decoded_family
            if decoded_topic and decoded_topic in PRODUCT_GROUPS[decoded_family]["topics"]:
                st.session_state.selected_topic = decoded_topic
                st.session_state.view = "topic"
            else:
                st.session_state.selected_topic = None
                st.session_state.view = "group"

    if str(st.query_params.get("admin","")).lower() in {"1","true","yes"}:
        if st.session_state.get("admin_authenticated"):
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
    if st.session_state.get("admin_authenticated"):
        render_admin()
    else:
        st.session_state.view = "home"
        st.rerun()
else:
    st.session_state.view = "home"
    st.rerun()

st.markdown(
    '<div class="footer">HPE Knowledge Base • AI-assisted technical knowledge search • '
    'Validate production procedures against current authoritative HPE documentation.</div>',
    unsafe_allow_html=True
)
