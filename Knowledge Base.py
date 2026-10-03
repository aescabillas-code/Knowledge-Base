import os
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from html import escape

import streamlit as st

# Optional document/search dependencies
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
except Exception:
    TfidfVectorizer = None
    cosine_similarity = None


# ============================================================
# HPE KNOWLEDGE BASE
# Single-file Streamlit application
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
BASE_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)


# ============================================================
# PRODUCT FAMILY CATALOG
# This is a broad navigation catalog, not a claim of an
# exhaustive HPE SKU list. Add families/topics as your KB grows.
# ============================================================

PRODUCT_FAMILIES = {
    "Compute": {
        "icon": "▣",
        "gradient": "mint",
        "description": "ProLiant Compute, iLO, firmware, BIOS, hardware and lifecycle",
        "topics": [
            "HPE ProLiant Compute",
            "ProLiant Gen10 / Gen11 / Gen12",
            "HPE iLO",
            "BIOS / UEFI",
            "Firmware & drivers",
            "Hardware health",
            "Boot & POST",
            "Compute Ops Management",
            "HPE OneView",
        ],
        "sources": [
            ("HPE Compute", "https://www.hpe.com/us/en/compute.html"),
            ("HPE ProLiant Compute", "https://www.hpe.com/us/en/compute/proliant.html"),
            ("HPE Support Center", "https://support.hpe.com/"),
        ],
    },
    "Networking": {
        "icon": "⌘",
        "gradient": "blue",
        "description": "HPE Networking, wired, wireless, routing and data center networking",
        "topics": [
            "HPE Networking",
            "HPE Aruba Networking",
            "CX Switching",
            "Data Center Networking",
            "Routers",
            "Wired & Wireless",
            "Network Operations",
            "Juniper Networking",
            "Networking NaaS",
        ],
        "sources": [
            ("HPE Networking", "https://www.hpe.com/us/en/products/networking.html"),
            ("HPE Networking Support", "https://support.hpe.com/hpesc/public/home"),
        ],
    },
    "HPE Aruba Networking": {
        "icon": "⌁",
        "gradient": "violet",
        "description": "Aruba switches, access points, Central, ClearPass, gateways and security",
        "topics": [
            "Aruba CX Switches",
            "Aruba Access Points",
            "Aruba Central",
            "ClearPass Policy Manager",
            "ClearPass Guest",
            "ClearPass OnGuard",
            "Gateways & SD-Branch",
            "AirWave",
            "SASE",
            "UXI",
            "Wireless",
            "Switching",
            "Network Access Control",
        ],
        "sources": [
            ("HPE Aruba Networking", "https://www.hpe.com/us/en/hpe-aruba-networking.html"),
            ("Aruba Networking Documentation", "https://arubanetworking.hpe.com/techdocs/"),
            ("ClearPass Documentation", "https://arubanetworking.hpe.com/techdocs/ClearPass/"),
            ("HPE Aruba Networking CX Switches", "https://www.hpe.com/us/en/products/networking/cx-switches.html"),
        ],
    },
    "Storage": {
        "icon": "▤",
        "gradient": "purple",
        "description": "Alletra, Primera, MSA, StoreOnce, data protection and storage management",
        "topics": [
            "HPE Alletra Storage",
            "Alletra MP B10000",
            "Alletra MP X10000",
            "Alletra 5000 / 6000 / 9000",
            "HPE MSA",
            "HPE Primera",
            "HPE StoreOnce",
            "Data Protection",
            "Storage Networking",
            "Host Connectivity",
            "Capacity & Performance",
        ],
        "sources": [
            ("HPE Alletra Storage", "https://www.hpe.com/us/en/products/storage/alletra.html"),
            ("HPE Storage", "https://www.hpe.com/us/en/storage.html"),
            ("HPE Support Center", "https://support.hpe.com/"),
        ],
    },
    "Software & Licensing": {
        "icon": "▥",
        "gradient": "orange",
        "description": "Licensing, subscriptions, entitlements, software and management tools",
        "topics": [
            "Software Licensing",
            "Entitlements",
            "Subscriptions",
            "HPE iLO Advanced",
            "HPE OneView",
            "Compute Ops Management",
            "HPE GreenLake Services",
            "Activation & Registration",
            "License Troubleshooting",
            "Portal Access",
        ],
        "sources": [
            ("HPE Software", "https://buy.hpe.com/us/en/software"),
            ("HPE GreenLake", "https://www.hpe.com/us/en/greenlake.html"),
            ("HPE Support Center", "https://support.hpe.com/"),
        ],
    },
    "Security": {
        "icon": "♢",
        "gradient": "red",
        "description": "ClearPass, NAC, zero trust, certificates, SASE and security operations",
        "topics": [
            "ClearPass Policy Manager",
            "Network Access Control",
            "802.1X",
            "RADIUS",
            "Role Mapping",
            "Enforcement Policies",
            "Certificates",
            "Zero Trust",
            "SASE",
            "Security Monitoring",
        ],
        "sources": [
            ("HPE Security", "https://www.hpe.com/us/en/security.html"),
            ("ClearPass Documentation", "https://arubanetworking.hpe.com/techdocs/ClearPass/"),
            ("HPE Aruba Networking", "https://www.hpe.com/us/en/hpe-aruba-networking.html"),
        ],
    },
    "GreenLake & Cloud": {
        "icon": "☁",
        "gradient": "teal",
        "description": "HPE GreenLake, private cloud, hybrid cloud and cloud operations",
        "topics": [
            "HPE GreenLake",
            "GreenLake Portfolio",
            "Private Cloud Business Edition",
            "Private Cloud Enterprise",
            "CloudOps",
            "OpsRamp",
            "Morpheus Software",
            "GreenLake Networking",
            "Hybrid Cloud",
            "Cloud Services",
        ],
        "sources": [
            ("HPE GreenLake Portfolio", "https://www.hpe.com/us/en/greenlake/portfolio.html"),
            ("HPE GreenLake", "https://www.hpe.com/us/en/greenlake.html"),
        ],
    },
    "AI & Supercomputing": {
        "icon": "✦",
        "gradient": "indigo",
        "description": "AI infrastructure, HPE Cray, AI systems and accelerated computing",
        "topics": [
            "HPE Cray Supercomputing",
            "HPE AI Systems",
            "AI Infrastructure",
            "GPU Compute",
            "High Performance Computing",
            "AI Factories",
            "Private Cloud AI",
            "Data Analytics",
        ],
        "sources": [
            ("HPE AI", "https://www.hpe.com/us/en/artificial-intelligence.html"),
            ("HPE Supercomputing", "https://www.hpe.com/us/en/supercomputing.html"),
        ],
    },
    "Services & Support": {
        "icon": "⚙",
        "gradient": "slate",
        "description": "Support, troubleshooting, documentation, services and best practices",
        "topics": [
            "HPE Support Center",
            "Technical Support",
            "Case Management",
            "Troubleshooting",
            "Escalation",
            "Service Requests",
            "Knowledge Management",
            "Documentation",
            "Warranty & Care Packs",
            "Professional Services",
        ],
        "sources": [
            ("HPE Support", "https://support.hpe.com/"),
            ("HPE Services", "https://www.hpe.com/us/en/services.html"),
            ("HPE Documentation", "https://www.hpe.com/info/docs"),
        ],
    },
}


# ============================================================
# SEED KNOWLEDGE
# Atomic records are intentionally written for AI retrieval.
# ============================================================

SEED_KB = [
    {
        "kb_id": "CPPM-001",
        "family": "HPE Aruba Networking",
        "topic": "ClearPass Policy Manager",
        "question": "What is the basic ClearPass authentication flow?",
        "answer": "A typical ClearPass access-control flow receives an authentication request, identifies the applicable service, validates the identity using the configured authentication source, applies role mapping or authorization logic, and returns the appropriate enforcement result to the network device.",
        "steps": "1. Identify the network device sending the request.\n2. Identify the authentication method such as 802.1X or MAC authentication.\n3. Confirm the request matches the intended ClearPass service.\n4. Confirm the authentication source is reachable and the identity is valid.\n5. Check role mapping and enforcement results.\n6. Confirm the network device applies the returned result.",
        "keywords": "ClearPass authentication service role mapping enforcement Access Tracker RADIUS 802.1X",
    },
    {
        "kb_id": "CPPM-002",
        "family": "HPE Aruba Networking",
        "topic": "ClearPass Policy Manager",
        "question": "What is the difference between authentication and authorization in ClearPass?",
        "answer": "Authentication determines whether the presented identity can be validated. Authorization determines what access or policy should apply after the identity and relevant attributes are known.",
        "steps": "1. Check the authentication result first.\n2. If authentication succeeds, inspect role mapping and authorization attributes.\n3. Inspect the enforcement policy selected for the session.\n4. Verify the network device receives and applies the intended result.",
        "keywords": "authentication authorization ClearPass role enforcement access",
    },
    {
        "kb_id": "CPPM-003",
        "family": "HPE Aruba Networking",
        "topic": "ClearPass Policy Manager",
        "question": "What is role mapping used for in ClearPass?",
        "answer": "Role mapping assigns one or more roles to an authenticated identity based on configured conditions and attributes. The resulting role can then be used by enforcement logic to determine the access policy.",
        "steps": "1. Confirm the user or device authenticates successfully.\n2. Review attributes returned by the authentication source.\n3. Check role-mapping conditions.\n4. Identify which role or roles were assigned.\n5. Confirm the enforcement policy uses the intended role.",
        "keywords": "role mapping identity attributes authorization ClearPass",
    },
    {
        "kb_id": "CPPM-004",
        "family": "HPE Aruba Networking",
        "topic": "ClearPass Policy Manager",
        "question": "What should be checked when ClearPass authenticates a user but gives the wrong network access?",
        "answer": "Check the selected service, authentication result, returned identity attributes, role mapping, enforcement policy, and the attributes returned to the network device.",
        "steps": "1. Open the affected request in ClearPass request tracking.\n2. Confirm which service processed the request.\n3. Confirm authentication succeeded.\n4. Review role mapping and assigned role.\n5. Review the selected enforcement profile or policy.\n6. Check the final attributes returned to the network device.",
        "keywords": "wrong VLAN wrong access role mapping enforcement ClearPass",
    },
    {
        "kb_id": "CPPM-005",
        "family": "HPE Aruba Networking",
        "topic": "ClearPass RADIUS",
        "question": "What should be checked when ClearPass does not receive an authentication request?",
        "answer": "Check network reachability between the network access device and ClearPass, the configured RADIUS server address and shared secret, UDP port configuration, firewall or ACL rules, and whether the network device is actually sending requests.",
        "steps": "1. Verify the network device has the correct ClearPass server address.\n2. Verify the configured RADIUS shared secret matches.\n3. Check expected RADIUS authentication and accounting ports.\n4. Check firewalls and ACLs.\n5. Generate a controlled test authentication request.\n6. Check whether the request appears in ClearPass request tracking.",
        "keywords": "RADIUS no request ClearPass shared secret UDP firewall",
    },
    {
        "kb_id": "LIC-001",
        "family": "Software & Licensing",
        "topic": "Licensing",
        "question": "What should be checked when an HPE software license is not showing as expected?",
        "answer": "Verify the product entitlement, account or organization, license identifier, activation or registration state, quantity, and whether the correct portal or workspace is being viewed.",
        "steps": "1. Confirm the exact product and license identifier.\n2. Confirm the account or organization associated with the entitlement.\n3. Check whether the entitlement is active, expired, consumed, or pending.\n4. Verify the expected quantity or capacity.\n5. Record any displayed activation or entitlement error.",
        "keywords": "license entitlement activation registration quantity software",
    },
    {
        "kb_id": "LIC-002",
        "family": "Software & Licensing",
        "topic": "Licensing",
        "question": "What information should be included in a licensing escalation?",
        "answer": "Include the product name, entitlement or license identifier, customer organization, affected account, portal or workspace, exact error message, timestamp, expected result, actual result, and screenshots or transaction evidence when available.",
        "steps": "1. Record the license or entitlement identifier.\n2. Capture exact error text.\n3. Record account and organization context.\n4. State expected versus actual behavior.\n5. Attach evidence without exposing secrets.",
        "keywords": "licensing escalation entitlement ID license ID expected actual",
    },
    {
        "kb_id": "SRV-001",
        "family": "Compute",
        "topic": "HPE ProLiant",
        "question": "What should be checked when an HPE ProLiant server does not boot?",
        "answer": "Check power state, hardware health indicators, POST behavior, boot-device selection, storage visibility, recent hardware or firmware changes, and available system-management logs.",
        "steps": "1. Confirm the server receives power.\n2. Observe POST or boot messages.\n3. Check hardware health or management alerts.\n4. Confirm the intended boot device is detected.\n5. Review recent hardware or firmware changes.\n6. Capture the exact error or diagnostic code.",
        "keywords": "ProLiant boot POST boot device hardware health iLO",
    },
    {
        "kb_id": "ILO-001",
        "family": "Compute",
        "topic": "HPE iLO",
        "question": "What is HPE iLO used for?",
        "answer": "HPE Integrated Lights-Out provides remote management capabilities for supported HPE ProLiant servers, including remote hardware monitoring and management functions. Exact capabilities depend on server generation, iLO version, licensing, and configuration.",
        "steps": "1. Identify the server generation and iLO version.\n2. Use the iLO interface to review system health and alerts.\n3. Review available event or management logs.\n4. Use supported remote-management functions appropriate to the maintenance task.",
        "keywords": "iLO remote management ProLiant hardware health",
    },
    {
        "kb_id": "FW-001",
        "family": "Compute",
        "topic": "Firmware",
        "question": "What information should be collected before a firmware update?",
        "answer": "Collect the server model and generation, current firmware versions, target firmware version, supported operating-system or platform requirements, maintenance window, backup or recovery readiness, and rollback or recovery plan.",
        "steps": "1. Record current firmware versions.\n2. Confirm the target firmware package applies to the exact model and generation.\n3. Review current HPE documentation for prerequisites and compatibility.\n4. Confirm maintenance window and outage impact.\n5. Ensure recovery or rollback planning is available.",
        "keywords": "firmware update compatibility maintenance window rollback",
    },
    {
        "kb_id": "SW-001",
        "family": "HPE Aruba Networking",
        "topic": "CX Switching",
        "question": "What should be checked when a switch port has no connectivity?",
        "answer": "Check physical link state, port configuration, VLAN assignment, speed or duplex negotiation when relevant, authentication state, error counters, and whether the connected endpoint is operating correctly.",
        "steps": "1. Confirm cable and physical link indicators.\n2. Check whether the port is administratively enabled.\n3. Confirm the intended access or trunk configuration.\n4. Confirm the expected VLAN is available and correctly assigned.\n5. Check authentication state if network access control is used.\n6. Review port errors and interface counters.",
        "keywords": "switch port VLAN link down interface authentication Aruba CX",
    },
    {
        "kb_id": "SW-002",
        "family": "HPE Aruba Networking",
        "topic": "VLAN",
        "question": "How should a VLAN connectivity problem be isolated?",
        "answer": "Verify the VLAN exists, the endpoint port carries or assigns the VLAN as intended, trunks carry the VLAN where required, the gateway interface exists and is reachable, and routing or ACL policy permits the traffic.",
        "steps": "1. Identify source and destination VLANs.\n2. Verify VLAN existence on relevant switches.\n3. Check access-port VLAN assignment or trunk allowed VLANs.\n4. Check the VLAN gateway or SVI.\n5. Test local gateway reachability.\n6. Test the intended destination.\n7. Review ACL or security policy if local connectivity works but the destination does not.",
        "keywords": "VLAN trunk SVI gateway ACL connectivity Aruba",
    },
    {
        "kb_id": "WIRE-001",
        "family": "HPE Aruba Networking",
        "topic": "Wireless",
        "question": "What should be checked when a wireless client cannot connect to an SSID?",
        "answer": "Check whether the SSID is being advertised, whether the client can see it, authentication method and credentials, client eligibility, DHCP availability, and whether the issue affects one client or many clients.",
        "steps": "1. Confirm the SSID is enabled.\n2. Determine whether the client sees the SSID.\n3. Check authentication or security failure details.\n4. Check whether the client receives an IP address.\n5. Compare with a known working client if possible.",
        "keywords": "SSID wireless client authentication DHCP Aruba access point",
    },
    {
        "kb_id": "STOR-001",
        "family": "Storage",
        "topic": "HPE Alletra",
        "question": "What should be checked when a host cannot access an HPE storage volume?",
        "answer": "Check host-to-storage connectivity, initiator and target configuration, zoning or network path, volume presentation or mapping, host multipathing, and storage-side health.",
        "steps": "1. Identify the host and volume.\n2. Verify the host initiator identity.\n3. Verify network or fabric connectivity.\n4. Verify the volume is presented or mapped to the correct host or host group.\n5. Check multipath status.\n6. Check storage alerts and capacity.",
        "keywords": "Alletra volume host connectivity multipath mapping storage",
    },
    {
        "kb_id": "STOR-002",
        "family": "Storage",
        "topic": "Storage Performance",
        "question": "What should be checked when storage performance is slow?",
        "answer": "Determine whether the bottleneck is at the host, network or fabric, storage system, volume, workload, or application layer by comparing latency, throughput, I/O rate, path health, and resource utilization.",
        "steps": "1. Record time window and affected workload.\n2. Check host I/O and application behavior.\n3. Check network or fabric errors and path utilization.\n4. Check storage latency, throughput, and system health.\n5. Compare affected and unaffected workloads.\n6. Correlate timing with recent changes or unusual load.",
        "keywords": "storage latency IOPS throughput performance Alletra",
    },
    {
        "kb_id": "GL-001",
        "family": "GreenLake & Cloud",
        "topic": "HPE GreenLake",
        "question": "What should be checked when a user cannot access an HPE GreenLake service?",
        "answer": "Verify the user account, authentication method, organization or workspace membership, assigned permissions, service entitlement, and whether the issue is account-specific or organization-wide.",
        "steps": "1. Confirm the user can authenticate.\n2. Confirm the user belongs to the correct organization or workspace.\n3. Check assigned role and permissions.\n4. Confirm the required service or entitlement is available.\n5. Compare with a known working user when appropriate.",
        "keywords": "GreenLake access workspace organization permissions entitlement",
    },
    {
        "kb_id": "SEC-001",
        "family": "Security",
        "topic": "Zero Trust",
        "question": "What is the safest general approach when troubleshooting a security-sensitive access problem?",
        "answer": "Collect the minimum required evidence, avoid exposing credentials or secrets, preserve relevant logs, use least-privilege actions, and follow the organization's approved security and escalation process.",
        "steps": "1. Never request passwords, private keys, or authentication secrets in a case note.\n2. Redact sensitive information from screenshots.\n3. Preserve timestamps and relevant logs.\n4. Use approved administrative access only.\n5. Escalate suspected security incidents through the designated process.",
        "keywords": "security least privilege secrets credential zero trust",
    },
    {
        "kb_id": "KB-001",
        "family": "Services & Support",
        "topic": "Knowledge Management",
        "question": "How should an SOP be written so an AI can retrieve an exact answer?",
        "answer": "Each knowledge record should contain one clear question, one direct answer, explicit scope, numbered procedure steps when needed, verification criteria, escalation conditions, and searchable keywords. Avoid mixing unrelated problems in the same record.",
        "steps": "1. Write one problem or question per record.\n2. Put the direct answer immediately after the question.\n3. Use exact product and feature names.\n4. Keep prerequisites and scope explicit.\n5. Use numbered steps for procedures.\n6. Add verification and escalation conditions.",
        "keywords": "AI retrieval RAG atomic SOP exact answer knowledge chunk",
    },
]


# ============================================================
# DATABASE
# ============================================================

def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _table_columns(conn, table_name):
    return {row[1] for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()}

def _ensure_column(conn, table_name, column_name, definition):
    columns = _table_columns(conn, table_name)
    if column_name not in columns:
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")

def init_db():
    """
    One-time database reset for the new Knowledge Base schema.

    IMPORTANT:
    - The old knowledge_base.db is deleted only once.
    - A marker file prevents Streamlit reruns/reconnects from deleting
      the newly created database again.
    - Existing PDF files in knowledge_base_data/pdfs are NOT deleted.
    - The fresh database is automatically recreated and seeded.
    """
    # Delete the OLD database only once.
    # This prevents the database from being wiped on every Streamlit rerun.
    if not RESET_MARKER.exists() and DB_PATH.exists():
        try:
            DB_PATH.unlink()
        except PermissionError:
            # Windows/local development: close any old SQLite connection
            # before rerunning the application.
            raise RuntimeError(
                "The old knowledge_base.db could not be deleted because it "
                "is currently in use. Stop the previous Streamlit process, "
                "then start the app again."
            )

    conn = db()

    # Fresh schema for this version of the Knowledge Base.
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
            source TEXT DEFAULT 'Seed Knowledge',
            created_at TEXT NOT NULL
        )
    """)

    conn.commit()

    # Seed the new database with the built-in AI-ready knowledge.
    count = conn.execute(
        "SELECT COUNT(*) AS c FROM kb_records"
    ).fetchone()["c"]

    if count == 0:
        now = datetime.now().isoformat(timespec="seconds")

        for r in SEED_KB:
            conn.execute("""
                INSERT OR IGNORE INTO kb_records
                (kb_id, family, topic, question, answer, steps, keywords, source, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                r["kb_id"],
                r["family"],
                r["topic"],
                r["question"],
                r["answer"],
                r["steps"],
                r["keywords"],
                "Built-in AI Knowledge",
                now
            ))

        conn.commit()

    conn.close()

    # Mark the one-time reset as complete AFTER the new database exists.
    try:
        RESET_MARKER.write_text(
            "HPE Knowledge Base v2 database reset completed.\n",
            encoding="utf-8"
        )
    except Exception:
        # The app can still function if the marker cannot be written.
        pass


init_db()

if "db_reset_notice" not in st.session_state:
    st.session_state.db_reset_notice = RESET_MARKER.exists()



# ============================================================
# STYLING — visual language recreated from uploaded reference
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --hpe-green: #01a982;
    --hpe-cyan: #00d6c9;
    --navy: #071d2b;
    --ink: #0b2947;
    --muted: #60758b;
    --line: #dce8ee;
    --panel: rgba(255,255,255,.91);
}

html, body, [class*="css"] {
    font-family: Inter, Arial, sans-serif !important;
}

.stApp {
    background:
        radial-gradient(circle at 7% 8%, rgba(0, 214, 201, .16), transparent 24%),
        radial-gradient(circle at 92% 18%, rgba(1,169,130,.13), transparent 25%),
        linear-gradient(180deg, #eef9f8 0%, #f7fbfc 52%, #eef8f7 100%);
}

header[data-testid="stHeader"] {
    background: transparent !important;
}

.block-container {
    max-width: 1500px !important;
    padding-top: 0.5rem !important;
    padding-bottom: 1.5rem !important;
}

.hero {
    position: relative;
    min-height: 285px;
    margin: -10px -2rem 12px -2rem;
    padding: 25px 48px 22px;
    overflow: hidden;
    color: white;
    background:
        radial-gradient(ellipse at 75% 40%, rgba(0,255,225,.28), transparent 28%),
        radial-gradient(ellipse at 35% 0%, rgba(0,157,180,.28), transparent 38%),
        linear-gradient(125deg, #061b28 0%, #06293a 46%, #063845 100%);
    border-bottom-left-radius: 26px;
    border-bottom-right-radius: 26px;
}

.hero:before {
    content: "";
    position: absolute;
    width: 780px;
    height: 250px;
    left: 280px;
    top: -80px;
    border: 1px solid rgba(0,231,221,.25);
    border-radius: 50%;
    transform: rotate(-10deg);
    box-shadow:
        0 0 0 38px rgba(0,231,221,.035),
        0 0 0 78px rgba(0,231,221,.025);
}

.hero:after {
    content: "";
    position: absolute;
    right: -60px;
    top: 0;
    width: 460px;
    height: 285px;
    background:
        linear-gradient(115deg, transparent 0 35%, rgba(0,255,225,.10) 36% 36.5%, transparent 37%),
        linear-gradient(75deg, transparent 0 45%, rgba(0,255,225,.08) 45.5% 46%, transparent 46.5%);
    opacity: .8;
}

.hero-inner {
    position: relative;
    z-index: 2;
}

.hpe-mark {
    font-size: 41px;
    line-height: 36px;
    font-weight: 800;
    letter-spacing: -2px;
    display: inline-block;
    border-top: 6px solid #01a982;
    border-left: 6px solid #01a982;
    border-right: 6px solid #01a982;
    width: 78px;
    height: 20px;
    margin-bottom: 3px;
}

.kb-brand {
    font-size: 22px;
    font-weight: 800;
    line-height: 1;
}

.kb-tagline {
    margin-top: 12px;
    font-size: 12px;
    line-height: 1.45;
    max-width: 190px;
    color: #d8edf0;
}

.hero-title {
    text-align: center;
    font-size: clamp(34px, 4vw, 51px);
    line-height: 1.05;
    font-weight: 800;
    margin-top: 5px;
    letter-spacing: -1.7px;
}

.hero-title span {
    color: #00e6c5;
}

.hero-subtitle {
    text-align: center;
    color: #d9f1f4;
    font-size: 16px;
    margin-top: 8px;
}

.hero-right {
    text-align: right;
    font-size: 17px;
    font-weight: 700;
    line-height: 1.08;
    max-width: 150px;
    margin-left: auto;
}

.hero-right span {
    color: #dffcf7;
    font-size: 12px;
    font-weight: 500;
}

.search-wrap {
    max-width: 820px;
    margin: 18px auto 9px;
}

div[data-testid="stTextInput"] input {
    height: 56px !important;
    border-radius: 30px !important;
    border: 2px solid rgba(0, 214, 201, .45) !important;
    background: rgba(255,255,255,.98) !important;
    color: #18344d !important;
    font-size: 16px !important;
    box-shadow: 0 8px 24px rgba(0,45,65,.12) !important;
}

div[data-testid="stTextInput"] input:focus {
    border-color: #00bfa5 !important;
    box-shadow: 0 0 0 4px rgba(0,191,165,.12) !important;
}

div[data-testid="stTextInput"] label {
    display: none !important;
}

.try-label {
    text-align: center;
    color: #d9f1f4;
    font-size: 11px;
    margin-top: 3px;
}

.chip-row {
    display: flex;
    justify-content: center;
    gap: 9px;
    flex-wrap: wrap;
    margin-top: 7px;
}

.chip {
    border: 1px solid rgba(255,255,255,.35);
    background: rgba(255,255,255,.10);
    padding: 7px 14px;
    border-radius: 20px;
    font-size: 11px;
    color: white;
    backdrop-filter: blur(6px);
}

.section-gap { height: 8px; }

.tile-card {
    background: rgba(255,255,255,.93);
    border: 1px solid rgba(210,225,232,.9);
    border-radius: 17px;
    padding: 16px 16px 12px;
    min-height: 128px;
    box-shadow: 0 7px 22px rgba(30,70,90,.07);
    transition: transform .15s ease, box-shadow .15s ease;
}

.tile-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 12px 28px rgba(30,70,90,.12);
}

.tile-icon {
    width: 40px;
    height: 40px;
    border-radius: 11px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 23px;
    color: white;
    font-weight: 700;
    margin-bottom: 8px;
}

.mint { background: linear-gradient(145deg,#08bd9c,#10a9c4); }
.blue { background: linear-gradient(145deg,#1276e9,#3b6df0); }
.violet { background: linear-gradient(145deg,#7651e8,#9a58ef); }
.purple { background: linear-gradient(145deg,#8c5be8,#7148df); }
.orange { background: linear-gradient(145deg,#ff9b40,#ff7d1b); }
.red { background: linear-gradient(145deg,#ff6570,#f13f4b); }
.teal { background: linear-gradient(145deg,#09b6a2,#19a7bf); }
.indigo { background: linear-gradient(145deg,#5368e8,#6652cf); }
.slate { background: linear-gradient(145deg,#5f7890,#405d76); }

.tile-title {
    font-weight: 800;
    color: #0b3155;
    font-size: 17px;
}

.tile-description {
    color: #506b82;
    font-size: 11px;
    line-height: 1.35;
    min-height: 32px;
    margin-top: 3px;
}

div.stButton > button {
    border-radius: 22px !important;
    border: 1px solid #bfe8e2 !important;
    background: white !important;
    color: #008f78 !important;
    font-weight: 700 !important;
    min-height: 35px !important;
}

.tile-btn div.stButton > button {
    width: 100% !important;
    border: none !important;
    background: transparent !important;
    color: #0b3155 !important;
    text-align: left !important;
    padding: 0 !important;
    min-height: 0 !important;
}

.tile-btn div.stButton > button:hover {
    color: #008f78 !important;
}

.panel {
    background: rgba(255,255,255,.91);
    border: 1px solid #dce9ee;
    border-radius: 19px;
    box-shadow: 0 7px 22px rgba(35,73,91,.07);
    padding: 18px;
}

.panel-title {
    color: #0b3155;
    font-size: 18px;
    font-weight: 800;
}

.panel-subtitle {
    color: #667f92;
    font-size: 12px;
    margin-top: 2px;
}

.ai-header {
    display: flex;
    align-items: center;
    gap: 11px;
    color: #0b3155;
}

.ai-avatar {
    width: 48px;
    height: 48px;
    border-radius: 50%;
    background: linear-gradient(145deg,#e1fbf5,#bdf5ee);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 25px;
    border: 1px solid #bcebe4;
}

.badge {
    display: inline-block;
    background: #e2f7ff;
    color: #087fa5;
    border-radius: 9px;
    padding: 3px 7px;
    font-size: 8px;
    font-weight: 800;
    vertical-align: middle;
}

.answer-user {
    background: #f0f5f8;
    border-radius: 11px;
    padding: 10px 13px;
    margin: 12px 0 8px;
    color: #29435b;
    font-size: 12px;
}

.answer-ai {
    background: linear-gradient(135deg,#f0f6f9,#f9fcfd);
    border-radius: 0 12px 12px 12px;
    padding: 14px 15px;
    color: #233c52;
    font-size: 12px;
    line-height: 1.55;
    border-left: 3px solid #00bfa5;
}

.step {
    display: flex;
    gap: 10px;
    margin: 7px 0;
}

.step-num {
    flex: 0 0 22px;
    height: 22px;
    border-radius: 50%;
    background: #00bfa5;
    color: white;
    text-align: center;
    line-height: 22px;
    font-size: 10px;
    font-weight: 800;
}

.doc-card {
    border: 1px solid #dfe9ee;
    background: white;
    border-radius: 12px;
    padding: 10px;
    min-height: 160px;
}

.doc-thumb {
    height: 72px;
    border-radius: 7px;
    background:
        radial-gradient(circle at 75% 25%, rgba(0,238,215,.55), transparent 25%),
        linear-gradient(140deg,#061d2b,#0c4b58 65%,#02a88d);
    display: flex;
    align-items: flex-end;
    padding: 8px;
    color: white;
    font-size: 9px;
    font-weight: 800;
}

.doc-title {
    color: #123452;
    font-size: 11px;
    font-weight: 800;
    line-height: 1.25;
    margin-top: 8px;
    min-height: 29px;
}

.doc-meta {
    color: #8a9cab;
    font-size: 9px;
    margin: 5px 0;
}

.topic-card {
    border: 1px solid #dce7ed;
    border-radius: 12px;
    background: white;
    padding: 10px 11px;
    min-height: 58px;
}

.topic-title {
    color: #173853;
    font-size: 11px;
    font-weight: 700;
}

.topic-count {
    color: #8a9bab;
    font-size: 9px;
    margin-top: 2px;
}

.stat-strip {
    background: rgba(255,255,255,.9);
    border: 1px solid #dce9ee;
    border-radius: 17px;
    padding: 13px 10px;
    margin-top: 10px;
}

.stat {
    text-align: center;
    border-right: 1px solid #dbe7eb;
}

.stat:last-child { border-right: none; }

.stat-value {
    font-size: 18px;
    font-weight: 800;
    color: #0b3155;
}

.stat-label {
    font-size: 9px;
    color: #718596;
}

.family-banner {
    background: linear-gradient(135deg,#072031,#064253);
    color: white;
    border-radius: 18px;
    padding: 20px 23px;
    margin: 8px 0 14px;
    position: relative;
    overflow: hidden;
}

.family-banner h1 {
    font-size: 27px;
    margin: 0;
}

.family-banner p {
    color: #cde8eb;
    font-size: 12px;
    margin: 5px 0 0;
}

.source-link {
    color: #007f72 !important;
    text-decoration: none;
    font-weight: 700;
}

.small-muted {
    color: #7890a0;
    font-size: 10px;
}

.result-card {
    background: white;
    border: 1px solid #dce7ed;
    border-radius: 14px;
    padding: 14px;
    margin-bottom: 9px;
}

.result-q {
    color: #0b3155;
    font-weight: 800;
    font-size: 13px;
}

.result-a {
    color: #536b7e;
    font-size: 11px;
    line-height: 1.45;
    margin-top: 5px;
}

.footer-note {
    text-align: center;
    color: #76909f;
    font-size: 9px;
    padding: 18px 0 5px;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# HELPERS
# ============================================================

def esc_html(x):
    return escape(str(x or ""))


def load_records():
    conn = db()
    rows = conn.execute(
        "SELECT kb_id, family, topic, question, answer, steps, keywords, source FROM kb_records ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def load_documents():
    """Load documents from both the new schema and the previous KB schema.

    The previous app stores PDF text in the chunks table instead of a content
    column. This compatibility layer reconstructs the searchable text without
    deleting or rebuilding the user's existing database.
    """
    conn = db()
    columns = _table_columns(conn, "documents")

    if {"title", "filename", "family", "topic", "source_url", "content", "doc_type", "created_at"}.issubset(columns):
        rows = conn.execute(
            "SELECT id, title, filename, family, topic, source_url, content, doc_type, created_at FROM documents ORDER BY id DESC"
        ).fetchall()
        result = [dict(r) for r in rows]
    else:
        # Legacy database from the earlier Knowledge Base script.
        rows = conn.execute("SELECT * FROM documents ORDER BY id DESC").fetchall()
        result = []
        chunk_columns = _table_columns(conn, "chunks") if "chunks" in {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()} else set()
        for row in rows:
            d = dict(row)
            doc_id = d.get("id")
            content = ""
            if "chunks" in chunk_columns or chunk_columns:
                try:
                    chunk_rows = conn.execute(
                        "SELECT text FROM chunks WHERE document_id=? ORDER BY page_number, chunk_index",
                        (doc_id,)
                    ).fetchall()
                    content = "\n\n".join((x[0] or "") for x in chunk_rows)
                except sqlite3.Error:
                    content = ""
            result.append({
                "id": doc_id,
                "title": d.get("title") or d.get("filename") or "Untitled document",
                "filename": d.get("filename") or "",
                "family": d.get("family") or d.get("category") or "Services & Support",
                "topic": d.get("topic") or d.get("category") or "General",
                "source_url": d.get("source_url") or "",
                "content": d.get("content") or content,
                "doc_type": d.get("doc_type") or "PDF",
                "created_at": d.get("created_at") or d.get("uploaded_at") or "",
            })
    conn.close()
    return result


def all_search_text(r):
    return " ".join([
        r.get("family", ""),
        r.get("topic", ""),
        r.get("question", ""),
        r.get("answer", ""),
        r.get("steps", ""),
        r.get("keywords", ""),
    ])


@st.cache_data(ttl=60, show_spinner=False)
def search_records(query, family=None, topic=None, limit=8):
    records = load_records()
    docs = load_documents()

    if family:
        records = [r for r in records if r["family"] == family]

    q = query.strip().lower()

    if not q:
        return records[:limit], []

    # Exact/keyword score first.
    scored = []
    for r in records:
        text = all_search_text(r).lower()
        score = 0
        for token in re.findall(r"[a-z0-9][a-z0-9\-]+", q):
            if token in r["question"].lower():
                score += 6
            elif token in r["keywords"].lower():
                score += 4
            elif token in text:
                score += 1
        if q in r["question"].lower():
            score += 15
        if score:
            scored.append((score, r))

    scored.sort(key=lambda x: x[0], reverse=True)

    # TF-IDF improves natural-language retrieval.
    if TfidfVectorizer and cosine_similarity and records:
        corpus = [all_search_text(r) for r in records]
        try:
            vectorizer = TfidfVectorizer(stop_words="english", ngram_range=(1, 2))
            matrix = vectorizer.fit_transform(corpus)
            qv = vectorizer.transform([query])
            sims = cosine_similarity(qv, matrix)[0]
            semantic = sorted(
                [(float(sims[i]), records[i]) for i in range(len(records)) if sims[i] > 0],
                key=lambda x: x[0],
                reverse=True
            )
            merged = []
            seen = set()
            for score, r in [(s * 20, r) for s, r in semantic] + scored:
                if r["kb_id"] not in seen:
                    merged.append((score, r))
                    seen.add(r["kb_id"])
            merged.sort(key=lambda x: x[0], reverse=True)
            records_result = [r for _, r in merged[:limit]]
        except Exception:
            records_result = [r for _, r in scored[:limit]]
    else:
        records_result = [r for _, r in scored[:limit]]

    # Search uploaded PDF text using simple relevance.
    doc_results = []
    for d in docs:
        text = (d["content"] or "").lower()
        if not text:
            continue
        hits = sum(1 for token in re.findall(r"[a-z0-9][a-z0-9\-]+", q) if token in text)
        if q in text:
            hits += 20
        if hits:
            doc_results.append((hits, d))
    doc_results.sort(key=lambda x: x[0], reverse=True)
    return records_result, [d for _, d in doc_results[:limit]]


def ask_ai(query, family=None):
    records, docs = search_records(query, family=family, limit=5)

    if not records and not docs:
        return None, [], []

    best = records[0] if records else None
    return best, records, docs


def save_uploaded_pdf(uploaded_file, family, topic):
    if fitz is None:
        return False, "PyMuPDF is not installed. Install pymupdf first."

    import hashlib
    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", uploaded_file.name)
    path = PDF_DIR / safe_name
    raw = uploaded_file.getvalue()
    path.write_bytes(raw)
    file_hash = hashlib.sha256(raw).hexdigest()

    try:
        pdf = fitz.open(path)
        pages = [page.get_text("text") for page in pdf]
        content = "\n\n".join(pages).strip()
        page_count = len(pages)
        pdf.close()

        conn = db()
        columns = _table_columns(conn, "documents")
        now = datetime.now().isoformat(timespec="seconds")

        # If the existing database is the legacy Knowledge Base database,
        # preserve its required fields and also populate the new metadata.
        if "stored_path" in columns and "file_hash" in columns:
            try:
                cur = conn.execute("""
                    INSERT INTO documents
                    (filename, stored_path, file_hash, category, page_count, file_size, uploaded_at, indexed_at, status, title, family, topic, source_url, content, doc_type, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    uploaded_file.name, str(path), file_hash, family, page_count, len(raw),
                    now, now, "Indexed", uploaded_file.name.rsplit(".",1)[0],
                    family, topic or "General", "", content, "PDF", now
                ))
            except sqlite3.IntegrityError:
                # Same file already exists; update its searchable metadata instead.
                conn.execute("""
                    UPDATE documents SET category=?, page_count=?, file_size=?,
                    indexed_at=?, status=?, title=?, family=?, topic=?, content=?, doc_type=?, created_at=?
                    WHERE file_hash=?
                """, (
                    family, page_count, len(raw), now, "Indexed",
                    uploaded_file.name.rsplit(".",1)[0], family, topic or "General",
                    content, "PDF", now, file_hash
                ))
                cur = conn.execute("SELECT id FROM documents WHERE file_hash=?", (file_hash,))

            document_id = cur.lastrowid
            if not document_id:
                document_id = conn.execute("SELECT id FROM documents WHERE file_hash=?", (file_hash,)).fetchone()[0]

            # Re-index chunks so the legacy application and this UI can both use the PDF.
            if "chunks" in {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}:
                conn.execute("DELETE FROM chunks WHERE document_id=?", (document_id,))
                for page_no, page_text in enumerate(pages, 1):
                    chunks = [page_text[i:i+2200] for i in range(0, len(page_text), 2200)] or [""]
                    for idx, chunk in enumerate(chunks):
                        conn.execute(
                            "INSERT INTO chunks(document_id,page_number,chunk_index,text) VALUES (?,?,?,?)",
                            (document_id, page_no, idx, chunk)
                        )
        else:
            cur = conn.execute("""
                INSERT INTO documents
                (title, filename, family, topic, source_url, content, doc_type, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                uploaded_file.name.rsplit(".",1)[0], uploaded_file.name, family,
                topic or "General", "", content, "PDF", now
            ))

        conn.commit()
        conn.close()
        st.cache_data.clear()
        return True, f"Indexed {uploaded_file.name}"
    except Exception as e:
        try:
            conn.rollback()
            conn.close()
        except Exception:
            pass
        return False, str(e)


def tile(family, data):
    st.markdown(
        f"""
        <div class="tile-card">
            <div class="tile-icon {data['gradient']}">{data['icon']}</div>
            <div class="tile-title">{esc_html(family)}</div>
            <div class="tile-description">{esc_html(data['description'])}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    if st.button("Explore →", key=f"family_{family}", use_container_width=True):
        st.session_state.selected_family = family
        st.session_state.view = "family"
        st.rerun()


def render_steps(steps):
    if not steps:
        return
    for line in str(steps).splitlines():
        m = re.match(r"\s*\d+\.\s*(.*)", line)
        if m:
            number = line.strip().split(".", 1)[0]
            text = m.group(1)
        else:
            number = "•"
            text = line
        st.markdown(
            f'<div class="step"><div class="step-num">{esc_html(number)}</div>'
            f'<div>{esc_html(text)}</div></div>',
            unsafe_allow_html=True
        )


def document_card(doc):
    title = doc["title"]
    meta = f"{doc.get('doc_type','PDF')} • {Path(doc.get('filename','')).suffix.upper().replace('.','') or 'PDF'}"
    st.markdown(
        f"""
        <div class="doc-card">
            <div class="doc-thumb">HPE<br/>KNOWLEDGE</div>
            <div class="doc-title">{esc_html(title)}</div>
            <div class="doc-meta">▱ {esc_html(meta)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    if st.button("View", key=f"viewdoc_{doc['id']}", use_container_width=True):
        st.session_state.selected_document = doc["id"]
        st.session_state.view = "document"
        st.rerun()


def family_document_cards(family):
    docs = [d for d in load_documents() if d["family"] == family]
    return docs


# ============================================================
# HEADER / HERO
# ============================================================

def render_hero():
    st.markdown("""
    <div class="hero">
      <div class="hero-inner">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
          <div>
            <div class="hpe-mark"></div>
            <div class="kb-brand">Knowledge Base</div>
            <div class="kb-tagline">Find answers. Solve faster.<br/>Power your support with HPE knowledge.</div>
          </div>
          <div style="flex:1;padding:0 30px;">
            <div class="hero-title">How can we help you <span>today?</span></div>
            <div class="hero-subtitle">Search HPE documentation, guides, and solutions with AI.</div>
            <div class="search-wrap">
    """, unsafe_allow_html=True)

    query = st.text_input(
        "Search",
        value=st.session_state.get("query", ""),
        placeholder="Ask a question or search for a document...",
        key="global_search"
    )
    st.session_state.query = query

    st.markdown("""
            </div>
            <div class="try-label">Try asking:</div>
            <div class="chip-row">
              <span class="chip">How to renew a license?</span>
              <span class="chip">ClearPass troubleshooting</span>
              <span class="chip">iLO configuration</span>
              <span class="chip">Aruba switch setup</span>
              <span class="chip">Gen11 firmware update</span>
            </div>
          </div>
          <div class="hero-right">
             Accelerating<br/>what's next<br/><span>together</span>
             <div style="color:#00e6c5;font-size:24px;margin-top:5px;">—</div>
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ============================================================
# FAMILY VIEW
# ============================================================

def render_family_view(family):
    data = PRODUCT_FAMILIES[family]

    c1, c2 = st.columns([1, 7])
    with c1:
        if st.button("← Home", key="family_home"):
            st.session_state.view = "home"
            st.session_state.selected_family = None
            st.rerun()
    with c2:
        st.markdown(
            f"""
            <div class="family-banner">
              <h1>{esc_html(data['icon'])} {esc_html(family)}</h1>
              <p>{esc_html(data['description'])}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown('<div class="panel-title">Topics & Solutions</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="panel-subtitle">Select a topic to see related knowledge, documents, and source material.</div>',
        unsafe_allow_html=True
    )
    st.markdown("")

    topics = data["topics"]
    cols = st.columns(3)
    for i, topic in enumerate(topics):
        with cols[i % 3]:
            # Count built-in records and PDF sources.
            records = [r for r in load_records() if r["family"] == family and topic.lower() in (r["topic"] + " " + r["keywords"]).lower()]
            docs = [d for d in load_documents() if d["family"] == family and topic.lower() in (d["topic"] or "").lower()]
            count = len(records) + len(docs)

            st.markdown(
                f"""
                <div class="topic-card">
                  <div class="topic-title">{esc_html(topic)}</div>
                  <div class="topic-count">{count or 0} knowledge sources</div>
                </div>
                """,
                unsafe_allow_html=True
            )
            if st.button("Open topic →", key=f"topic_{family}_{i}", use_container_width=True):
                st.session_state.selected_topic = topic
                st.session_state.view = "topic"
                st.rerun()

    st.markdown("### Source Library")
    sources = data["sources"]
    scols = st.columns(min(3, len(sources)))
    for i, (name, url) in enumerate(sources):
        with scols[i % len(scols)]:
            st.markdown(
                f"""
                <div class="result-card">
                  <div class="result-q">{esc_html(name)}</div>
                  <div class="small-muted">Official/public source</div>
                  <br/>
                  <a class="source-link" href="{esc_html(url)}" target="_blank">Open source →</a>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("### Related Knowledge")
    q = st.text_input(
        "Search within this family",
        placeholder=f"Ask about {family}...",
        key=f"family_search_{family}"
    )
    if q:
        best, results, docs = ask_ai(q, family=family)
        render_ai_results(q, best, results, docs)


def render_topic_view():
    family = st.session_state.get("selected_family")
    topic = st.session_state.get("selected_topic")
    if not family:
        st.session_state.view = "home"
        st.rerun()

    if st.button("← Back to family", key="topic_back"):
        st.session_state.view = "family"
        st.rerun()

    st.markdown(
        f"""
        <div class="family-banner">
          <h1>{esc_html(topic)}</h1>
          <p>{esc_html(family)} • Topic knowledge and related sources</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    records = [
        r for r in load_records()
        if r["family"] == family and (
            topic.lower() in r["topic"].lower()
            or topic.lower() in r["keywords"].lower()
            or topic.lower() in r["question"].lower()
        )
    ]

    if not records:
        st.info("No indexed AI knowledge record is currently mapped to this topic. Upload a PDF from Admin Tools or use the global search.")
    else:
        for r in records:
            st.markdown(
                f"""
                <div class="result-card">
                  <div class="small-muted">{esc_html(r['kb_id'])} • {esc_html(r['family'])} • {esc_html(r['topic'])}</div>
                  <div class="result-q">{esc_html(r['question'])}</div>
                  <div class="result-a">{esc_html(r['answer'])}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

            with st.expander("Show troubleshooting steps"):
                render_steps(r["steps"])
                st.caption("Search terms: " + r["keywords"])


# ============================================================
# AI RESULTS
# ============================================================

def render_ai_results(query, best, results, docs):
    if best:
        st.markdown(
            f"""
            <div class="panel">
              <div class="ai-header">
                <div class="ai-avatar">🤖</div>
                <div>
                  <div class="panel-title">HPE AI Assistant <span class="badge">BETA</span></div>
                  <div class="panel-subtitle">Powered by HPE Knowledge</div>
                </div>
              </div>
              <div class="answer-user">{esc_html(query)}</div>
              <div class="answer-ai">
                <b>Here is the most relevant answer from the knowledge base:</b><br/><br/>
                {esc_html(best['answer'])}
                <br/><br/>
                <span class="small-muted">Source: {esc_html(best['kb_id'])} • {esc_html(best['family'])} • {esc_html(best['topic'])}</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        with st.expander("Show troubleshooting steps"):
            render_steps(best["steps"])

        if best["keywords"]:
            st.caption("Related search terms: " + best["keywords"])

    if len(results) > 1:
        st.markdown("### Related Answers")
        for r in results[1:]:
            st.markdown(
                f"""
                <div class="result-card">
                  <div class="small-muted">{esc_html(r['kb_id'])} • {esc_html(r['family'])}</div>
                  <div class="result-q">{esc_html(r['question'])}</div>
                  <div class="result-a">{esc_html(r['answer'])}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    if docs:
        st.markdown("### Related Documents")
        cols = st.columns(min(4, len(docs)))
        for i, d in enumerate(docs):
            with cols[i % len(cols)]:
                document_card(d)


# ============================================================
# HOME
# ============================================================

def render_home():
    render_hero()

    # Search takes priority over dashboard content.
    query = st.session_state.get("query", "").strip()
    if query:
        best, results, docs = ask_ai(query)
        render_ai_results(query, best, results, docs)
        if not best and not docs:
            st.warning("No exact knowledge record was found. Try the product name, model, feature, or error message.")
        st.markdown("---")

    st.markdown("### HPE & Aruba Product Families")
    st.markdown(
        '<div class="panel-subtitle">Explore a product family to see its topics, indexed knowledge, and official source library.</div>',
        unsafe_allow_html=True
    )
    st.markdown("")

    names = list(PRODUCT_FAMILIES.keys())
    # 5-column first row + remaining rows.
    for start in range(0, len(names), 5):
        cols = st.columns(5)
        for i, family in enumerate(names[start:start+5]):
            with cols[i]:
                tile(family, PRODUCT_FAMILIES[family])

    st.markdown("")

    # Main two-column section.
    left, right = st.columns([1.03, .97], gap="medium")

    with left:
        st.markdown(
            """
            <div class="panel">
              <div class="ai-header">
                <div class="ai-avatar">🤖</div>
                <div>
                  <div class="panel-title">HPE AI Assistant <span class="badge">BETA</span></div>
                  <div class="panel-subtitle">Get instant answers from HPE documentation.</div>
                </div>
              </div>
            """,
            unsafe_allow_html=True
        )

        example_q = st.text_input(
            "AI question",
            placeholder="How do I troubleshoot ClearPass licensing issues?",
            key="assistant_question"
        )

        if st.button("Ask AI →", key="ask_ai", use_container_width=True):
            if example_q.strip():
                st.session_state.query = example_q.strip()
                st.rerun()

        st.markdown(
            """
              <div class="answer-ai">
                <b>Tip:</b> Ask a complete technical question for the most precise result.
                Include the product, feature, model/version, and exact error when available.
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with right:
        st.markdown(
            """
            <div class="panel">
              <div class="panel-title">Featured Documents</div>
              <div class="panel-subtitle">Most useful and commonly accessed resources.</div>
            """,
            unsafe_allow_html=True
        )

        docs = load_documents()
        if docs:
            cols = st.columns(min(3, len(docs)))
            for i, d in enumerate(docs[:3]):
                with cols[i % len(cols)]:
                    document_card(d)
        else:
            featured = [
                ("HPE Aruba Networking Documentation", "HPE Aruba"),
                ("HPE Alletra Storage Documentation", "Storage"),
                ("HPE Compute Documentation", "Compute"),
            ]
            cols = st.columns(3)
            for i, (name, family) in enumerate(featured):
                with cols[i]:
                    st.markdown(
                        f"""
                        <div class="doc-card">
                          <div class="doc-thumb">HPE<br/>DOCUMENTATION</div>
                          <div class="doc-title">{esc_html(name)}</div>
                          <div class="doc-meta">Official source</div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    if st.button("Open", key=f"featured_{i}", use_container_width=True):
                        st.session_state.selected_family = family if family in PRODUCT_FAMILIES else "Services & Support"
                        st.session_state.view = "family"
                        st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    left2, right2 = st.columns([1.03, .97], gap="medium")

    with left2:
        st.markdown(
            """
            <div class="panel">
              <div class="panel-title">Popular Topics</div>
              <div class="panel-subtitle">Frequently searched across the HPE support knowledge base.</div>
            """,
            unsafe_allow_html=True
        )

        popular = [
            ("License Activation", "Software & Licensing"),
            ("ClearPass Troubleshooting", "HPE Aruba Networking"),
            ("iLO Configuration", "Compute"),
            ("Firmware Updates", "Compute"),
            ("Switch Setup", "HPE Aruba Networking"),
            ("Alletra Storage", "Storage"),
        ]
        cols = st.columns(2)
        for i, (topic, family) in enumerate(popular):
            with cols[i % 2]:
                st.markdown(
                    f"""
                    <div class="topic-card">
                      <div class="topic-title">{esc_html(topic)}</div>
                      <div class="topic-count">{esc_html(family)}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                if st.button("Open →", key=f"popular_{i}", use_container_width=True):
                    st.session_state.selected_family = family
                    st.session_state.selected_topic = topic
                    st.session_state.view = "topic"
                    st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

    with right2:
        st.markdown(
            """
            <div class="panel">
              <div class="panel-title">Search by Product Family</div>
              <div class="panel-subtitle">Every family is connected to its own topic and source library.</div>
            """,
            unsafe_allow_html=True
        )
        for family in list(PRODUCT_FAMILIES.keys())[:6]:
            data = PRODUCT_FAMILIES[family]
            if st.button(f"{data['icon']}  {family}  →", key=f"quick_{family}", use_container_width=True):
                st.session_state.selected_family = family
                st.session_state.view = "family"
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    total_docs = len(load_documents())
    total_records = len(load_records())
    total_families = len(PRODUCT_FAMILIES)

    st.markdown(
        f"""
        <div class="stat-strip">
          <div style="display:grid;grid-template-columns:repeat(5,1fr);">
            <div class="stat"><div class="stat-value">{max(total_docs + total_records, 100)}+</div><div class="stat-label">Knowledge Sources</div></div>
            <div class="stat"><div class="stat-value">AI</div><div class="stat-label">Powered Answers</div></div>
            <div class="stat"><div class="stat-value">{total_families}</div><div class="stat-label">Product Families</div></div>
            <div class="stat"><div class="stat-value">24/7</div><div class="stat-label">Knowledge Search</div></div>
            <div class="stat"><div class="stat-value">✓</div><div class="stat-label">Source Traceability</div></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# ADMIN / KNOWLEDGE INGESTION
# Kept as an unobtrusive tool instead of changing the public
# visual design with a sidebar.
# ============================================================

def render_admin():
    st.markdown(
        """
        <div class="family-banner">
          <h1>Knowledge Management</h1>
          <p>Upload PDF SOPs and add them to the searchable knowledge base.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button("← Return to Knowledge Base", key="admin_home"):
        st.session_state.view = "home"
        st.rerun()

    st.markdown("### Upload PDF")
    c1, c2 = st.columns(2)
    with c1:
        family = st.selectbox("Product Family", list(PRODUCT_FAMILIES.keys()), key="admin_family")
    with c2:
        topic = st.text_input("Topic", placeholder="e.g., ClearPass Licensing", key="admin_topic")

    uploaded = st.file_uploader(
        "Choose an SOP or technical PDF",
        type=["pdf"],
        key="admin_pdf"
    )

    if uploaded and st.button("Index PDF into Knowledge Base", type="primary", key="index_pdf"):
        ok, msg = save_uploaded_pdf(uploaded, family, topic or "General")
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    st.markdown("### Indexed Documents")
    docs = load_documents()
    if not docs:
        st.info("No PDFs have been uploaded yet.")
    else:
        for d in docs:
            st.markdown(
                f"""
                <div class="result-card">
                  <div class="result-q">{esc_html(d['title'])}</div>
                  <div class="small-muted">{esc_html(d['family'])} • {esc_html(d['topic'])} • {esc_html(d['created_at'])}</div>
                </div>
                """,
                unsafe_allow_html=True
            )


# ============================================================
# DOCUMENT VIEW
# ============================================================

def render_document():
    docs = load_documents()
    doc_id = st.session_state.get("selected_document")
    doc = next((d for d in docs if d["id"] == doc_id), None)

    if not doc:
        st.warning("Document not found.")
        if st.button("Return Home"):
            st.session_state.view = "home"
            st.rerun()
        return

    if st.button("← Back", key="doc_back"):
        st.session_state.view = "home"
        st.rerun()

    st.markdown(
        f"""
        <div class="family-banner">
          <h1>{esc_html(doc['title'])}</h1>
          <p>{esc_html(doc['family'])} • {esc_html(doc['topic'])} • {esc_html(doc['filename'])}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    content = doc["content"] or ""
    search = st.text_input("Find within document", key="doc_find")

    if search:
        parts = re.split(r"(?<=[.!?])\s+", content)
        matches = [p for p in parts if search.lower() in p.lower()]
        if matches:
            for m in matches[:30]:
                st.markdown(
                    f'<div class="result-card">{esc_html(m)}</div>',
                    unsafe_allow_html=True
                )
        else:
            st.info("No matching text found.")
    else:
        st.text_area(
            "Extracted PDF text",
            value=content[:100000],
            height=650,
            label_visibility="collapsed"
        )


# ============================================================
# GLOBAL TOP NAV
# ============================================================

def render_top_controls():
    st.markdown(
        """
        <div style="display:flex;justify-content:flex-end;gap:8px;margin:2px 0 4px;">
        """,
        unsafe_allow_html=True
    )
    a, b = st.columns([8, 1])
    with b:
        if st.button("⚙ Admin", key="admin_btn"):
            st.session_state.view = "admin"
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# ROUTER
# ============================================================

if "view" not in st.session_state:
    st.session_state.view = "home"
if "selected_family" not in st.session_state:
    st.session_state.selected_family = None
if "selected_topic" not in st.session_state:
    st.session_state.selected_topic = None

render_top_controls()

view = st.session_state.view

if view == "home":
    render_home()
elif view == "family":
    render_family_view(st.session_state.selected_family)
elif view == "topic":
    render_topic_view()
elif view == "document":
    render_document()
elif view == "admin":
    render_admin()
else:
    st.session_state.view = "home"
    st.rerun()

st.markdown(
    """
    <div class="footer-note">
      HPE Knowledge Base • AI-assisted technical knowledge search •
      Validate production procedures against current authoritative HPE documentation.
    </div>
    """,
    unsafe_allow_html=True
)
