import streamlit as st
import google.generativeai as genai
import json, json5, re, os, time, tempfile, requests
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Credit Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=DM+Serif+Display&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #0D0D14;
    color: #E2DDD4;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #13131E;
    border-right: 1px solid #1E1E2E;
}
[data-testid="stSidebar"] * { color: #C5BFB3 !important; }

/* Headers */
h1 { font-family: 'DM Serif Display', serif !important; color: #C9A84C !important; letter-spacing: -0.5px; }
h2, h3 { color: #C9A84C !important; font-weight: 600 !important; }

/* Inputs */
input, textarea, select, [data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
    background-color: #1A1A28 !important;
    border: 1px solid #2A2A3E !important;
    color: #E2DDD4 !important;
    border-radius: 6px !important;
}
input:focus, textarea:focus {
    border-color: #C9A84C !important;
    box-shadow: 0 0 0 2px rgba(201,168,76,0.15) !important;
}

/* Labels */
label { color: #9D9890 !important; font-size: 0.82rem !important; font-weight: 500 !important; letter-spacing: 0.04em; text-transform: uppercase; }

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #C9A84C, #A8872E) !important;
    color: #0D0D14 !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 0.6rem 1.5rem !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.02em;
    transition: opacity 0.2s;
}
.stButton > button:hover { opacity: 0.88 !important; }

/* Download button */
.stDownloadButton > button {
    background: #1A1A28 !important;
    color: #C9A84C !important;
    border: 1px solid #C9A84C !important;
    font-weight: 600 !important;
    border-radius: 6px !important;
}

/* Expander */
[data-testid="stExpander"] {
    background-color: #13131E !important;
    border: 1px solid #1E1E2E !important;
    border-radius: 8px !important;
}

/* Tabs */
[data-testid="stTabs"] button {
    color: #9D9890 !important;
    font-weight: 500 !important;
    border-bottom: 2px solid transparent !important;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #C9A84C !important;
    border-bottom-color: #C9A84C !important;
}

/* Radio */
[data-testid="stRadio"] label { text-transform: none !important; font-size: 0.9rem !important; }

/* Section card */
.section-card {
    background: #13131E;
    border: 1px solid #1E1E2E;
    border-radius: 10px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
}
.section-label {
    color: #C9A84C;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 0.8rem;
}
.url-tag {
    display: inline-block;
    background: #1E1E2E;
    color: #C9A84C;
    font-size: 0.72rem;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 4px;
    margin-bottom: 4px;
    letter-spacing: 0.06em;
}
.status-ok { color: #4CAF50; font-size: 0.85rem; }
.status-warn { color: #FF9800; font-size: 0.85rem; }
.header-bar {
    background: linear-gradient(135deg, #13131E, #1A1A28);
    border-left: 3px solid #C9A84C;
    padding: 1.2rem 1.5rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    st.markdown("---")

    api_key = st.text_input("Gemini API Key", type="password",
                            placeholder="AIza...",
                            help="Get free key at aistudio.google.com")

    st.markdown("---")
    st.markdown("**Entity Type**")
    entity_type = st.radio("", [
        "NBFC / IFC",
        "Bank / SFB",
        "PSU Finance",
        "Apex DFI"
    ], label_visibility="collapsed")

    st.markdown("---")
    st.markdown("**Prepared By**")
    prepared_by = st.text_input("Analyst Name", placeholder="Your name", label_visibility="collapsed")
    st.markdown("**Reviewed By**")
    reviewed_by = st.text_input("CIO / Manager Name", placeholder="CIO name", label_visibility="collapsed")

    st.markdown("---")
    st.caption("v3.0 · Gemini File API · Accurate extraction")


# ─────────────────────────────────────────────
# MAIN HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1 style="margin:0;font-size:1.8rem;">Credit Report Generator</h1>
    <p style="margin:0.3rem 0 0;color:#9D9890;font-size:0.88rem;">
        AI-powered · Gemini File API · Accurate financial extraction
    </p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# ISSUER DETAILS
# ─────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Issuer Details</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    issuer_name = st.text_input("Issuer Name", placeholder="e.g. Power Finance Corporation Ltd")
with col2:
    industry = st.text_input("Industry (NIC)", placeholder="e.g. Financial Services")
with col3:
    sector = st.text_input("Sector", placeholder="e.g. NBFC – Infrastructure Finance")

col4, col5 = st.columns(2)
with col4:
    review_period = st.text_input("Review Period", placeholder="e.g. H1FY26 (April – September 2025)")
with col5:
    business_house = st.text_input("Business House / Group", placeholder="e.g. Government of India")

st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# INVESTMENTS TABLE
# ─────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Investment Details</div>', unsafe_allow_html=True)
st.caption("Add each bond / NCD held. Click + Add Bond to add more rows.")

if "bonds" not in st.session_state:
    st.session_state.bonds = [{"security": "", "yield": "", "agency": "", "rating": "", "fv": "", "sh": "", "ph": ""}]

def add_bond():
    st.session_state.bonds.append({"security": "", "yield": "", "agency": "", "rating": "", "fv": "", "sh": "", "ph": ""})

def remove_bond(i):
    if len(st.session_state.bonds) > 1:
        st.session_state.bonds.pop(i)

header_cols = st.columns([3, 1.2, 1.2, 1.2, 1, 1, 1, 0.5])
for col, label in zip(header_cols, ["Security / Maturity Date", "Yield (%)", "Rating Agency", "Rating", "FV (Cr)", "SH (Cr)", "PH (Cr)", ""]):
    col.markdown(f"<small style='color:#9D9890;font-weight:600;text-transform:uppercase;letter-spacing:0.05em'>{label}</small>", unsafe_allow_html=True)

for i, bond in enumerate(st.session_state.bonds):
    cols = st.columns([3, 1.2, 1.2, 1.2, 1, 1, 1, 0.5])
    st.session_state.bonds[i]["security"] = cols[0].text_input("sec", value=bond["security"], key=f"sec_{i}", label_visibility="collapsed", placeholder="NCD DD.MM.YYYY")
    st.session_state.bonds[i]["yield"]    = cols[1].text_input("yld", value=bond["yield"],    key=f"yld_{i}", label_visibility="collapsed", placeholder="7.50")
    st.session_state.bonds[i]["agency"]   = cols[2].text_input("agc", value=bond["agency"],   key=f"agc_{i}", label_visibility="collapsed", placeholder="CRISIL")
    st.session_state.bonds[i]["rating"]   = cols[3].text_input("rtg", value=bond["rating"],   key=f"rtg_{i}", label_visibility="collapsed", placeholder="AAA")
    st.session_state.bonds[i]["fv"]       = cols[4].text_input("fv",  value=bond["fv"],       key=f"fv_{i}",  label_visibility="collapsed", placeholder="25")
    st.session_state.bonds[i]["sh"]       = cols[5].text_input("sh",  value=bond["sh"],       key=f"sh_{i}",  label_visibility="collapsed", placeholder="")
    st.session_state.bonds[i]["ph"]       = cols[6].text_input("ph",  value=bond["ph"],       key=f"ph_{i}",  label_visibility="collapsed", placeholder="25")
    if cols[7].button("✕", key=f"del_{i}"):
        remove_bond(i)
        st.rerun()

st.button("+ Add Bond", on_click=add_bond)
st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SOURCE DOCUMENTS — FILE API APPROACH
# ─────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Source Documents</div>', unsafe_allow_html=True)
st.caption("Paste direct PDF URLs (from NSE/BSE filings) **or** upload files. PDFs are sent directly to Gemini for accurate visual reading — no text extraction.")

doc_tab, upload_tab = st.tabs(["📎 Paste URLs", "⬆️ Upload Files"])

with doc_tab:
    st.markdown("**H1FY26** — April to September 2025")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<span class="url-tag">Q1FY26 · Apr–Jun 2025</span>', unsafe_allow_html=True)
        q1fy26_url = st.text_input("q1fy26", placeholder="Direct PDF link from NSE/BSE", label_visibility="collapsed", key="q1fy26_url")
    with c2:
        st.markdown('<span class="url-tag">Q2FY26 · Jul–Sep 2025</span>', unsafe_allow_html=True)
        q2fy26_url = st.text_input("q2fy26", placeholder="Direct PDF link from NSE/BSE", label_visibility="collapsed", key="q2fy26_url")

    st.markdown("**H1FY25** — April to September 2024")
    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<span class="url-tag">Q1FY25 · Apr–Jun 2024</span>', unsafe_allow_html=True)
        q1fy25_url = st.text_input("q1fy25", placeholder="Direct PDF link from NSE/BSE", label_visibility="collapsed", key="q1fy25_url")
    with c4:
        st.markdown('<span class="url-tag">Q2FY25 · Jul–Sep 2024</span>', unsafe_allow_html=True)
        q2fy25_url = st.text_input("q2fy25", placeholder="Direct PDF link from NSE/BSE", label_visibility="collapsed", key="q2fy25_url")

    st.markdown("**Annual Reports**")
    c5, c6 = st.columns(2)
    with c5:
        st.markdown('<span class="url-tag">FY25 Annual Report · 31.03.2025</span>', unsafe_allow_html=True)
        fy25_url = st.text_input("fy25", placeholder="Direct PDF link from NSE/BSE", label_visibility="collapsed", key="fy25_url")
    with c6:
        st.markdown('<span class="url-tag">FY24 Annual Report · 31.03.2024</span>', unsafe_allow_html=True)
        fy24_url = st.text_input("fy24", placeholder="Direct PDF link from NSE/BSE", label_visibility="collapsed", key="fy24_url")

with upload_tab:
    st.caption("Upload PDFs directly. Hold Ctrl to select multiple files at once.")
    uploaded_files = st.file_uploader(
        "Upload quarterly results and annual reports",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    if uploaded_files:
        for f in uploaded_files:
            st.markdown(f'<span class="status-ok">✓ {f.name}</span>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# HELPER: Download PDF from URL
# ─────────────────────────────────────────────
def download_pdf(url: str, label: str) -> tuple[bytes | None, str]:
    """Download PDF from URL. Returns (bytes, filename) or (None, error)."""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/pdf,*/*"
        }
        r = requests.get(url.strip(), headers=headers, timeout=30)
        r.raise_for_status()
        ct = r.headers.get("Content-Type", "")
        if "pdf" not in ct.lower() and not url.strip().lower().endswith(".pdf"):
            # Try to find PDF link on page
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(r.text, "html.parser")
            pdf_links = []
            for a in soup.find_all("a", href=True):
                href = a["href"]
                text = a.get_text().lower()
                if ".pdf" in href.lower():
                    score = 0
                    for kw in ["annual report", "quarterly", "financial result", "q1", "q2", "q3", "q4"]:
                        if kw in text or kw in href.lower():
                            score += 3
                    for kw in ["csr", "sustainability", "agm notice", "postal ballot"]:
                        if kw in text or kw in href.lower():
                            score -= 5
                    pdf_links.append((score, href))
            if pdf_links:
                pdf_links.sort(reverse=True)
                best = pdf_links[0][1]
                if not best.startswith("http"):
                    from urllib.parse import urljoin
                    best = urljoin(url, best)
                r2 = requests.get(best, headers=headers, timeout=30)
                r2.raise_for_status()
                return r2.content, f"{label}.pdf"
            return None, f"No PDF found at {url}"
        return r.content, f"{label}.pdf"
    except Exception as e:
        return None, str(e)


# ─────────────────────────────────────────────
# HELPER: Upload file to Gemini File API
# ─────────────────────────────────────────────
def upload_to_gemini(pdf_bytes: bytes, filename: str) -> object | None:
    """Upload PDF bytes to Gemini File API. Returns file object."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        gfile = genai.upload_file(tmp_path, mime_type="application/pdf", display_name=filename)
        # Wait for processing
        for _ in range(20):
            gfile = genai.get_file(gfile.name)
            if gfile.state.name == "ACTIVE":
                return gfile
            time.sleep(2)
        return None
    except Exception as e:
        st.warning(f"⚠️ Could not upload {filename}: {e}")
        return None
    finally:
        os.unlink(tmp_path)


# ─────────────────────────────────────────────
# HELPER: Build Gemini prompt
# ─────────────────────────────────────────────
def build_prompt(issuer_name, entity_type, industry, sector, review_period,
                 business_house, prepared_by, reviewed_by, bonds,
                 doc_labels: list[str]) -> str:

    bonds_str = "\n".join([
        f"  - Security: {b['security']} | Yield: {b['yield']}% | Agency: {b['agency']} | Rating: {b['rating']} | FV: {b['fv']} Cr | SH: {b['sh']} | PH: {b['ph']}"
        for b in bonds if b.get("security")
    ])

    entity_rules = {
        "NBFC / IFC": """
INCLUDE: Total Income, Net Interest Income (NII), Profit After Tax (PAT), Total Assets, 
Loans & Advances / Loan Book, Borrowings, Net Worth, GNPA / Gross Stage 3 (%), 
NNPA / Net Stage 3 (%), CRAR / CAR (%), Debt/Equity Ratio, EPS (Rs), NIM (%), ROA (%), ROE (%)
EXCLUDE: Deposits, CASA%, Tier I CAR% (unless stated)
NOTE: GNPA may be called "Gross Stage 3", "Gross Impaired Assets", or "Asset Quality Ratio" — treat all as equivalent""",

        "Bank / SFB": """
INCLUDE: Total Income, Net Interest Income (NII), Profit After Tax (PAT), Total Assets, 
Gross Advances, Deposits, CASA Ratio (%), GNPA (%), NNPA (%), CRAR / CAR (%), Tier I CAR (%),
NIM (%), ROA (%), ROE (%)
EXCLUDE: Debt/Equity (not applicable for banks)""",

        "PSU Finance": """
INCLUDE: Total Income, Net Interest Income (NII), Profit After Tax (PAT), Total Assets,
Loans & Advances, Borrowings, Net Worth, GNPA / Gross Stage 3 (%), NNPA / Net Stage 3 (%),
CRAR (%), Debt/Equity, EPS (Rs), ROA (%), ROE (%)
EXCLUDE: Deposits
NOTE: Comment on Government of India ownership and policy mandate""",

        "Apex DFI": """
INCLUDE: Total Income, Net Interest Income (NII), Profit After Tax (PAT), Total Assets,
Loan Disbursements, Loan Outstanding / Advances, Borrowings, Net Worth, GNPA (%),
NNPA (%), CRAR (%), ROA (%), Government support status
EXCLUDE: Deposits from public, CASA%
NOTE: Mention GoI ownership, developmental mandate, and off-balance sheet exposure if stated"""
    }

    doc_note = f"The following {len(doc_labels)} PDF documents have been uploaded for you to read directly:\n" + \
               "\n".join([f"  • {l}" for l in doc_labels]) if doc_labels else \
               "No documents provided — use your knowledge of this company's publicly disclosed financials."

    return f"""You are a senior credit analyst at an Indian health insurance company preparing a formal credit report for the investment committee. 

ISSUER: {issuer_name}
ENTITY TYPE: {entity_type}
INDUSTRY: {industry}
SECTOR: {sector}
REVIEW PERIOD: {review_period}
BUSINESS HOUSE: {business_house}
PREPARED BY: {prepared_by}
REVIEWED BY: {reviewed_by}

BONDS / NCDs HELD:
{bonds_str}

SOURCE DOCUMENTS:
{doc_note}

=== FINANCIAL METRICS TO EXTRACT (based on entity type) ===
{entity_rules.get(entity_type, entity_rules["NBFC / IFC"])}

=== PERIOD COLUMNS REQUIRED ===
The financial table MUST have exactly 4 columns:
  Col 1: H1FY26  (April – September 2025)
  Col 2: H1FY25  (April – September 2024)
  Col 3: 31.03.2025 (Full Year FY25)
  Col 4: 31.03.2024 (Full Year FY24)

=== H1 CALCULATION RULES — CRITICAL ===
FLOW ITEMS — ADD Q1 + Q2 (income statement items):
  Total Income, Net Interest Income, PAT, Provisions, Fee Income
  H1FY26 = Q1FY26 value + Q2FY26 value
  H1FY25 = Q1FY25 value + Q2FY25 value

STOCK ITEMS — USE Q2 VALUE ONLY (balance sheet at period end):
  Total Assets, Loans & Advances, Borrowings, Net Worth, Deposits
  H1FY26 = Q2FY26 value (as at 30 Sep 2025)
  H1FY25 = Q2FY25 value (as at 30 Sep 2024)

RATIO ITEMS — USE Q2 VALUE (point-in-time):
  GNPA%, NNPA%, CRAR%, NIM%, ROA%, ROE%, Debt/Equity, CASA%
  H1FY26 = Q2FY26 ratio
  H1FY25 = Q2FY25 ratio

ANNUAL FIGURES: Use directly from FY25 Annual Report for 31.03.2025, FY24 Annual Report for 31.03.2024.

=== CALCULATION FORMULAS ===
When figures are not directly stated, calculate them:
- NII = Interest Income – Interest Expense
- GNPA% = Gross NPA (or Gross Stage 3) / Gross Advances × 100
- NNPA% = Net NPA (or Net Stage 3) / Net Advances × 100
- NIM% = Annualised NII / Average Interest-Earning Assets × 100  [For H1: use NII × 2]
- ROA% = Annualised PAT / Average Total Assets × 100              [For H1: use PAT × 2]
- ROE% = Annualised PAT / Average Net Worth × 100                 [For H1: use PAT × 2]
- Debt/Equity = Total Borrowings / Net Worth
- EPS = PAT (annualised) / Shares Outstanding
- PCR = Cumulative Provisions / Gross NPA × 100
- Credit Cost% = Annualised Provisions / Average Loan Book × 100

UNIT: All monetary values in Rs. Crores. Auto-convert: Lakhs ÷ 100 = Crores, Millions ÷ 10 = Crores.
STANDALONE ONLY: Always use standalone (not consolidated) figures.
NEVER output N/A — if a figure cannot be found or calculated, state "Not disclosed" or your best estimate.

=== OUTPUT FORMAT ===
Respond ONLY with valid JSON. No markdown, no explanation, no preamble.

{{
  "company_profile": "4-6 sentence formal paragraph about the company — history, ownership, business model, market position, regulatory standing",
  "financial_table": [
    {{"metric": "Total Income (Rs Cr)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "Net Interest Income (Rs Cr)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "Profit After Tax (Rs Cr)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "Total Assets (Rs Cr)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "Loans & Advances / Loan Book (Rs Cr)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "Borrowings (Rs Cr)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "Net Worth (Rs Cr)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "GNPA / Gross Stage 3 (%)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "NNPA / Net Stage 3 (%)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "CRAR / CAR (%)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "NIM (%)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "ROA (%)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "ROE (%)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}},
    {{"metric": "Debt/Equity (x)", "h1fy26": "...", "h1fy25": "...", "mar25": "...", "mar24": "..."}}
  ],
  "profitability": "2-3 sentences with specific numbers on income growth, NIM trend, PAT growth, ROA/ROE",
  "asset_quality": "2-3 sentences on GNPA/NNPA trend, Stage 3 movement, PCR, credit cost",
  "capitalisation": "2-3 sentences on CRAR adequacy, Tier I ratio, net worth growth, leverage",
  "liquidity": "2-3 sentences on borrowing profile, maturity mix, ALM, debt/equity, access to markets",
  "recommendation": "Paragraph starting with 'Keeping in view...' — summarise the investment case and recommend continuation/review of exposure"
}}"""


# ─────────────────────────────────────────────
# HELPER: Parse Gemini JSON response
# ─────────────────────────────────────────────
def parse_response(raw: str) -> dict:
    clean = re.sub(r"```json|```", "", raw).strip()
    match = re.search(r"\{[\s\S]*\}", clean)
    if match:
        clean = match.group(0)
    try:
        return json.loads(clean)
    except Exception:
        try:
            return json5.loads(clean)
        except Exception as e:
            raise ValueError(f"Could not parse Gemini response: {e}\n\nRaw (first 500 chars):\n{raw[:500]}")


# ─────────────────────────────────────────────
# HELPER: Build Word document
# ─────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "2A2A5A")
        tblBorders.append(border)
    tblPr.append(tblBorders)


def build_docx(data: dict, issuer_name, entity_type, industry, sector,
               review_period, business_house, prepared_by, reviewed_by, bonds) -> bytes:
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    # ── HEADER ──
    header_table = doc.add_table(rows=1, cols=1)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hc = header_table.rows[0].cells[0]
    set_cell_bg(hc, "1A1A2E")
    hp = hc.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = hp.add_run(issuer_name.upper())
    run.font.name = "Calibri"
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    hp2 = hc.add_paragraph()
    hp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = hp2.add_run("CREDIT REVIEW & ANALYSIS")
    r2.font.name = "Calibri"
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
    r2.font.bold = True

    doc.add_paragraph()

    # ── COMPANY PROFILE ──
    cp_heading = doc.add_paragraph()
    ch = cp_heading.add_run("COMPANY PROFILE")
    ch.font.name = "Calibri"
    ch.font.size = Pt(10)
    ch.font.bold = True
    ch.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)

    cp_para = doc.add_paragraph()
    cp_para.paragraph_format.space_after = Pt(8)
    cr = cp_para.add_run(data.get("company_profile", ""))
    cr.font.name = "Calibri"
    cr.font.size = Pt(9.5)

    # ── ISSUER INFORMATION TABLE ──
    doc.add_paragraph()
    issuer_heading = doc.add_paragraph()
    ih = issuer_heading.add_run("ISSUER INFORMATION")
    ih.font.name = "Calibri"
    ih.font.size = Pt(10)
    ih.font.bold = True
    ih.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)

    info_table = doc.add_table(rows=6, cols=2)
    add_border(info_table)
    info_rows = [
        ("Issuer", issuer_name),
        ("Industry (NIC)", industry),
        ("Sector", sector),
        ("Review Period", review_period),
        ("Prepared By", prepared_by),
        ("Reviewed By", reviewed_by),
    ]
    for i, (key, val) in enumerate(info_rows):
        row = info_table.rows[i]
        set_cell_bg(row.cells[0], "EFF6FF" if i % 2 == 0 else "FFFFFF")
        set_cell_bg(row.cells[1], "EFF6FF" if i % 2 == 0 else "FFFFFF")
        k_run = row.cells[0].paragraphs[0].add_run(key)
        k_run.font.name = "Calibri"
        k_run.font.size = Pt(9)
        k_run.font.bold = True
        v_run = row.cells[1].paragraphs[0].add_run(val)
        v_run.font.name = "Calibri"
        v_run.font.size = Pt(9)

    # ── INVESTMENTS TABLE ──
    doc.add_paragraph()
    inv_heading = doc.add_paragraph()
    ivh = inv_heading.add_run("BRIEF DETAIL OF INVESTMENT UNDER REVIEW")
    ivh.font.name = "Calibri"
    ivh.font.size = Pt(10)
    ivh.font.bold = True
    ivh.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)

    inv_cols = ["Security / Maturity Date", "Yield (%)", "Rating Agency", "Rating",
                "Business House", "F.V. (Cr)", "SH (Cr)", "PH (Cr)"]
    inv_table = doc.add_table(rows=1 + len([b for b in bonds if b.get("security")]), cols=8)
    add_border(inv_table)

    for j, col_name in enumerate(inv_cols):
        hcell = inv_table.rows[0].cells[j]
        set_cell_bg(hcell, "1A1A2E")
        hr = hcell.paragraphs[0].add_run(col_name)
        hr.font.name = "Calibri"
        hr.font.size = Pt(8)
        hr.font.bold = True
        hr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        hcell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    row_idx = 1
    for bond in bonds:
        if not bond.get("security"):
            continue
        row = inv_table.rows[row_idx]
        bg = "EFF6FF" if row_idx % 2 == 1 else "FFFFFF"
        vals = [bond["security"], bond["yield"] + "%", bond["agency"],
                bond["rating"], business_house, bond["fv"], bond["sh"], bond["ph"]]
        for j, val in enumerate(vals):
            set_cell_bg(row.cells[j], bg)
            r = row.cells[j].paragraphs[0].add_run(val or "—")
            r.font.name = "Calibri"
            r.font.size = Pt(8.5)
            row.cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_idx += 1

    # ── FINANCIAL STRENGTH TABLE ──
    doc.add_paragraph()
    fs_heading = doc.add_paragraph()
    fsh = fs_heading.add_run("FINANCIAL STRENGTH")
    fsh.font.name = "Calibri"
    fsh.font.size = Pt(10)
    fsh.font.bold = True
    fsh.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)

    fin_data = data.get("financial_table", [])
    fin_table = doc.add_table(rows=1 + len(fin_data), cols=5)
    add_border(fin_table)

    col_headers = ["Particulars", "H1FY26", "H1FY25", "31.03.2025", "31.03.2024"]
    for j, hdr in enumerate(col_headers):
        hcell = fin_table.rows[0].cells[j]
        set_cell_bg(hcell, "1A1A2E")
        hr = hcell.paragraphs[0].add_run(hdr)
        hr.font.name = "Calibri"
        hr.font.size = Pt(9)
        hr.font.bold = True
        hr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        hcell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, row_data in enumerate(fin_data):
        row = fin_table.rows[i + 1]
        bg = "EFF6FF" if i % 2 == 0 else "FFFFFF"
        vals = [
            row_data.get("metric", ""),
            row_data.get("h1fy26", "—"),
            row_data.get("h1fy25", "—"),
            row_data.get("mar25",  "—"),
            row_data.get("mar24",  "—"),
        ]
        for j, val in enumerate(vals):
            set_cell_bg(row.cells[j], bg)
            r = row.cells[j].paragraphs[0].add_run(str(val))
            r.font.name = "Calibri"
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
            else:
                row.cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── COMMENTS ──
    doc.add_paragraph()
    comments_heading = doc.add_paragraph()
    comh = comments_heading.add_run("COMMENTS")
    comh.font.name = "Calibri"
    comh.font.size = Pt(10)
    comh.font.bold = True
    comh.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)

    for label, key in [
        ("Profitability", "profitability"),
        ("Asset Quality", "asset_quality"),
        ("Capitalisation", "capitalisation"),
        ("Liquidity", "liquidity"),
    ]:
        subh_para = doc.add_paragraph()
        subh = subh_para.add_run(label)
        subh.font.name = "Calibri"
        subh.font.size = Pt(9.5)
        subh.font.bold = True
        subh.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)

        body_para = doc.add_paragraph()
        body_para.paragraph_format.space_after = Pt(6)
        br = body_para.add_run(data.get(key, ""))
        br.font.name = "Calibri"
        br.font.size = Pt(9.5)

    # ── RECOMMENDATION ──
    doc.add_paragraph()
    rec_table = doc.add_table(rows=1, cols=1)
    add_border(rec_table)
    rc = rec_table.rows[0].cells[0]
    set_cell_bg(rc, "F0FDF4")

    rh = rc.paragraphs[0].add_run("RECOMMENDATION")
    rh.font.name = "Calibri"
    rh.font.size = Pt(9.5)
    rh.font.bold = True
    rh.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    rp = rc.add_paragraph()
    rr = rp.add_run(data.get("recommendation", ""))
    rr.font.name = "Calibri"
    rr.font.size = Pt(9.5)

    # ── FOOTER ──
    doc.add_paragraph()
    footer_table = doc.add_table(rows=1, cols=2)
    add_border(footer_table)
    set_cell_bg(footer_table.rows[0].cells[0], "EFF6FF")
    set_cell_bg(footer_table.rows[0].cells[1], "EFF6FF")

    fm_run = footer_table.rows[0].cells[0].paragraphs[0].add_run(f"Fund Manager: {prepared_by}")
    fm_run.font.name = "Calibri"
    fm_run.font.size = Pt(8.5)
    fm_run.font.bold = True

    cio_para = footer_table.rows[0].cells[1].paragraphs[0]
    cio_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    cio_run = cio_para.add_run(f"CIO: {reviewed_by}")
    cio_run.font.name = "Calibri"
    cio_run.font.size = Pt(8.5)
    cio_run.font.bold = True

    # Save to bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# GENERATE BUTTON
# ─────────────────────────────────────────────
st.markdown("---")
col_gen, col_hint = st.columns([1, 3])
with col_gen:
    generate = st.button("⚡ Generate Report", use_container_width=True)
with col_hint:
    st.caption("PDFs are uploaded directly to Gemini — it reads tables visually, like a human analyst would.")

if generate:
    # Validate
    errors = []
    if not api_key:
        errors.append("Gemini API key is required (sidebar)")
    if not issuer_name:
        errors.append("Issuer name is required")
    if not prepared_by:
        errors.append("Analyst name is required (sidebar)")
    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    # Configure Gemini
    genai.configure(api_key=api_key)

    with st.status("Generating report...", expanded=True) as status:

        # ── Step 1: Collect & upload PDFs ──
        st.write("📥 Collecting source documents...")
        uploaded_gemini_files = []
        doc_labels = []

        url_map = {
            "Q1FY26 (Apr–Jun 2025)": st.session_state.get("q1fy26_url", ""),
            "Q2FY26 (Jul–Sep 2025)": st.session_state.get("q2fy26_url", ""),
            "Q1FY25 (Apr–Jun 2024)": st.session_state.get("q1fy25_url", ""),
            "Q2FY25 (Jul–Sep 2024)": st.session_state.get("q2fy25_url", ""),
            "FY25 Annual Report":     st.session_state.get("fy25_url", ""),
            "FY24 Annual Report":     st.session_state.get("fy24_url", ""),
        }

        for label, url in url_map.items():
            if url and url.strip():
                st.write(f"  ⬇️  Downloading {label}...")
                pdf_bytes, fname = download_pdf(url.strip(), label.replace(" ", "_").replace("/", "-"))
                if pdf_bytes:
                    st.write(f"  ⬆️  Uploading {label} to Gemini...")
                    gfile = upload_to_gemini(pdf_bytes, fname)
                    if gfile:
                        uploaded_gemini_files.append(gfile)
                        doc_labels.append(label)
                        st.write(f"  ✅ {label} ready")
                    else:
                        st.write(f"  ⚠️  {label} upload timed out — skipping")
                else:
                    st.write(f"  ⚠️  Could not download {label}: {fname}")

        # Handle uploaded files
        if uploaded_files:
            for uf in uploaded_files:
                st.write(f"  ⬆️  Uploading {uf.name} to Gemini...")
                gfile = upload_to_gemini(uf.read(), uf.name)
                if gfile:
                    uploaded_gemini_files.append(gfile)
                    doc_labels.append(uf.name)
                    st.write(f"  ✅ {uf.name} ready")

        if not uploaded_gemini_files:
            st.write("  ℹ️  No documents uploaded — Gemini will use its training knowledge")

        # ── Step 2: Call Gemini ──
        st.write("🤖 Sending to Gemini for analysis...")

        prompt = build_prompt(
            issuer_name, entity_type, industry, sector, review_period,
            business_house, prepared_by, reviewed_by,
            st.session_state.bonds, doc_labels
        )

        model = genai.GenerativeModel(
            model_name="gemini-2.5-flash",
            generation_config={"temperature": 0.1, "max_output_tokens": 16000}
        )

        # Build content parts: all uploaded PDFs + prompt
        content_parts = [f for f in uploaded_gemini_files] + [prompt]

        for attempt in range(3):
            try:
                response = model.generate_content(content_parts)
                raw = response.text
                break
            except Exception as e:
                if attempt < 2:
                    st.write(f"  ⏳ Rate limit — waiting 60s (attempt {attempt+1}/3)...")
                    time.sleep(60)
                else:
                    st.error(f"Gemini error after 3 attempts: {e}")
                    st.stop()

        # ── Step 3: Parse ──
        st.write("📋 Parsing response...")
        try:
            report_data = parse_response(raw)
        except ValueError as e:
            st.error(str(e))
            st.stop()

        # ── Step 4: Build Word doc ──
        st.write("📝 Building Word document...")
        docx_bytes = build_docx(
            report_data, issuer_name, entity_type, industry, sector,
            review_period, business_house, prepared_by, reviewed_by,
            st.session_state.bonds
        )

        # Clean up Gemini files
        for gf in uploaded_gemini_files:
            try:
                genai.delete_file(gf.name)
            except Exception:
                pass

        status.update(label="✅ Report generated!", state="complete")

    # ── RESULTS ──
    st.markdown("---")
    st.markdown("### 📄 Report Ready")

    fname_out = f"{issuer_name.replace(' ', '_')}_Credit_Report_{review_period.replace(' ', '').replace('–', '-')}.docx"
    st.download_button(
        label=f"⬇️  Download {fname_out}",
        data=docx_bytes,
        file_name=fname_out,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

    # Preview
    with st.expander("👁️ Preview — Company Profile & Recommendation"):
        st.markdown(f"**{issuer_name}**")
        st.write(report_data.get("company_profile", ""))
        st.markdown("---")
        st.markdown("**Recommendation**")
        st.write(report_data.get("recommendation", ""))

    with st.expander("📊 Preview — Financial Table"):
        fin_rows = report_data.get("financial_table", [])
        if fin_rows:
            import pandas as pd
            df = pd.DataFrame([{
                "Metric": r["metric"],
                "H1FY26": r.get("h1fy26", "—"),
                "H1FY25": r.get("h1fy25", "—"),
                "31.03.2025": r.get("mar25", "—"),
                "31.03.2024": r.get("mar24", "—"),
            } for r in fin_rows])
            st.dataframe(df, use_container_width=True, hide_index=True)
