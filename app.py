import streamlit as st
import json, re, io, os, base64, requests, subprocess, sys
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Credit Report Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');

/* Global */
html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Background */
.stApp {
    background: #0f0f13;
    color: #e8e4dc;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #16161c;
    border-right: 1px solid #2a2a35;
}

[data-testid="stSidebar"] .stMarkdown h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 1.4rem;
    color: #e8e4dc;
    letter-spacing: -0.5px;
}

/* Headers */
h1 {
    font-family: 'DM Serif Display', serif !important;
    font-size: 2.4rem !important;
    color: #e8e4dc !important;
    letter-spacing: -1px !important;
    line-height: 1.1 !important;
}

h2, h3 {
    font-family: 'DM Serif Display', serif !important;
    color: #e8e4dc !important;
}

/* Labels */
label, .stTextInput label, .stSelectbox label, .stTextArea label {
    color: #9b9688 !important;
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    letter-spacing: 1px !important;
    text-transform: uppercase !important;
}

/* Inputs */
.stTextInput input, .stTextArea textarea {
    background: #1c1c24 !important;
    border: 1px solid #2a2a35 !important;
    border-radius: 8px !important;
    color: #e8e4dc !important;
    font-family: 'DM Sans', sans-serif !important;
}

.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #c9a84c !important;
    box-shadow: 0 0 0 2px rgba(201,168,76,0.15) !important;
}

/* Selectbox */
.stSelectbox select, [data-baseweb="select"] {
    background: #1c1c24 !important;
    border: 1px solid #2a2a35 !important;
    color: #e8e4dc !important;
    border-radius: 8px !important;
}

/* Buttons */
.stButton button {
    background: #c9a84c !important;
    color: #0f0f13 !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'DM Sans', sans-serif !important;
    padding: 0.6rem 1.5rem !important;
    transition: all 0.2s !important;
    letter-spacing: 0.5px !important;
}

.stButton button:hover {
    background: #e0bc62 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 20px rgba(201,168,76,0.3) !important;
}

/* Download button */
.stDownloadButton button {
    background: #1c1c24 !important;
    color: #c9a84c !important;
    border: 1px solid #c9a84c !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'DM Sans', sans-serif !important;
    padding: 0.6rem 1.5rem !important;
    transition: all 0.2s !important;
}

.stDownloadButton button:hover {
    background: #c9a84c !important;
    color: #0f0f13 !important;
}

/* Cards */
.report-card {
    background: #16161c;
    border: 1px solid #2a2a35;
    border-radius: 12px;
    padding: 24px;
    margin-bottom: 16px;
}

.stat-card {
    background: #1c1c24;
    border: 1px solid #2a2a35;
    border-radius: 10px;
    padding: 16px 20px;
    text-align: center;
}

.stat-value {
    font-family: 'DM Serif Display', serif;
    font-size: 1.8rem;
    color: #c9a84c;
}

.stat-label {
    font-size: 0.7rem;
    color: #9b9688;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 4px;
}

/* Gold accent */
.gold { color: #c9a84c; }
.muted { color: #9b9688; font-size: 0.85rem; }

/* Tag */
.tag {
    display: inline-block;
    background: rgba(201,168,76,0.15);
    color: #c9a84c;
    border: 1px solid rgba(201,168,76,0.3);
    border-radius: 4px;
    padding: 2px 8px;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
}

/* Divider */
hr {
    border-color: #2a2a35 !important;
    margin: 24px 0 !important;
}

/* Success/info boxes */
.stSuccess {
    background: rgba(34,197,94,0.1) !important;
    border: 1px solid rgba(34,197,94,0.3) !important;
    border-radius: 8px !important;
}

.stError {
    background: rgba(239,68,68,0.1) !important;
    border: 1px solid rgba(239,68,68,0.3) !important;
    border-radius: 8px !important;
}

/* Hide streamlit branding */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Number input */
.stNumberInput input {
    background: #1c1c24 !important;
    border: 1px solid #2a2a35 !important;
    color: #e8e4dc !important;
    border-radius: 8px !important;
}

/* Radio */
.stRadio label {
    color: #e8e4dc !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    font-size: 0.9rem !important;
    font-weight: 400 !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #1c1c24 !important;
    border: 1px dashed #2a2a35 !important;
    border-radius: 10px !important;
}
</style>
""", unsafe_allow_html=True)

# ── DOCX builder (same as notebook) ──────────────────────────
DARK_NAVY   = RGBColor(0x1A,0x1A,0x2E)
BLUE_ACCENT = RGBColor(0x25,0x63,0xEB)
WHITE       = RGBColor(0xFF,0xFF,0xFF)
DARK_TEXT   = RGBColor(0x1F,0x29,0x37)
MUTED_C     = RGBColor(0x47,0x55,0x69)

def _bg(cell, h):
    s = OxmlElement('w:shd')
    s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),h)
    cell._tc.get_or_add_tcPr().append(s)

def _borders(cell):
    b = OxmlElement('w:tcBorders')
    for side in ('top','bottom','left','right'):
        e = OxmlElement(f'w:{side}')
        e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'2')
        e.set(qn('w:space'),'0');    e.set(qn('w:color'),'CCCCCC')
        b.append(e)
    cell._tc.get_or_add_tcPr().append(b)

def _cw(table, widths):
    tg = OxmlElement('w:tblGrid')
    for w in widths:
        gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(w)); tg.append(gc)
    table._tbl.insert(1, tg)
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            tw = OxmlElement('w:tcW')
            tw.set(qn('w:w'),str(w)); tw.set(qn('w:type'),'dxa')
            cell._tc.get_or_add_tcPr().append(tw)

def _r(para, text, bold=False, size=10.5, color=None):
    run = para.add_run(str(text))
    run.bold = bold; run.font.name = 'Calibri'; run.font.size = Pt(size)
    if color: run.font.color.rgb = color
    return run

def _hd(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    _r(p, text.upper(), bold=True, size=9.5, color=BLUE_ACCENT)
    pBdr = OxmlElement('w:pBdr')
    btm  = OxmlElement('w:bottom')
    btm.set(qn('w:val'),'single'); btm.set(qn('w:sz'),'8')
    btm.set(qn('w:space'),'1');    btm.set(qn('w:color'),'2563EB')
    pBdr.append(btm); p._p.get_or_add_pPr().append(pBdr)

def build_docx(data):
    doc = Document()
    for section in doc.sections:
        section.page_height=Cm(29.7); section.page_width=Cm(21.0)
        section.left_margin=Inches(1); section.right_margin=Inches(1)
        section.top_margin=Inches(1);  section.bottom_margin=Inches(1)

    entity  = data['entityType']
    periods = data.get('financialPeriods',['H1FY26','H1FY25','31.03.2025','31.03.2024'])

    # Header
    for text, sz, fill, clr in [
        (data['issuerName'].upper(), 18, '1A1A2E', WHITE),
        ('Credit Review & Analysis', 12, '1A1A2E', RGBColor(0x94,0xA3,0xB8)),
        (' ', 4, '2563EB', WHITE)
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
        shd.set(qn('w:fill'),fill); p._p.get_or_add_pPr().append(shd)
        _r(p, text, bold=(sz==18), size=sz, color=clr)

    # Company Profile
    _hd(doc,'Company Profile')
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(8)
    _r(p, data.get('companyProfile',''), size=10.5, color=DARK_TEXT)

    # Issuer Info
    _hd(doc,'Issuer Information')
    t = doc.add_table(rows=6, cols=2); t.style='Table Grid'; _cw(t,[2700,6300])
    for i,(label,value) in enumerate([
        ('Issuer', data['issuerName']),
        ('Industry as per NIC code', data.get('industry','')),
        ('Sector', data.get('sector','')),
        ('Review Period', data.get('reviewPeriod','')),
        ('Prepared By', data['preparedBy']),
        ('Reviewed By', data['reviewedBy']),
    ]):
        lc,vc = t.rows[i].cells[0], t.rows[i].cells[1]
        _bg(lc,'F1F5F9'); _borders(lc); _bg(vc,'FFFFFF'); _borders(vc)
        for cell,txt,bold,clr in [(lc,label,True,MUTED_C),(vc,value,False,DARK_TEXT)]:
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            _r(p, txt, bold=bold, size=9.5, color=clr)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Investments
    _hd(doc,'Brief Detail of Investment under Review')
    p2 = doc.add_paragraph(); p2.paragraph_format.space_after = Pt(4)
    _r(p2,'(Rs. in crores)', size=9, color=MUTED_C)
    invs = [i for i in data.get('investments',[]) if i.get('security')]
    if invs:
        hdrs = ['Security / Date of Maturity','Holding Yield','Credit Rating Agency',
                'Credit Rating','Business House','Total Nominal Value (F.V)','SH','PH']
        ti = doc.add_table(rows=1+len(invs), cols=8); ti.style='Table Grid'
        ws = [2100,680,1100,700,1100,1100,680,680]
        total=9360; s=sum(ws); ws=[int(w*total/s) for w in ws]; ws[-1]+=total-sum(ws)
        _cw(ti,ws)
        for ci,h in enumerate(hdrs):
            cell=ti.rows[0].cells[ci]; _bg(cell,'1A1A2E'); _borders(cell)
            p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
            _r(p,h,bold=True,size=8,color=WHITE)
        for ri,inv in enumerate(invs):
            row=ti.rows[ri+1]; bgc='FFFFFF' if ri%2==0 else 'F8FAFC'
            for ci,v in enumerate([inv.get('security',''),inv.get('yield',''),inv.get('agency',''),
                                    inv.get('rating',''),inv.get('businessHouse',''),
                                    inv.get('fv',''),inv.get('sh',''),inv.get('ph','')]):
                cell=row.cells[ci]; _bg(cell,bgc); _borders(cell)
                p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
                _r(p,v or '',size=9,color=DARK_TEXT)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Financial Strength
    _hd(doc,'Financial Strength')
    p3=doc.add_paragraph(); p3.paragraph_format.space_after=Pt(4)
    _r(p3,'(Rs. in crores)',size=9,color=MUTED_C)
    fin = list(data.get('financialData',[]))
    for m in data.get('conditionalMetrics',[]):
        if entity in m.get('includeFor',[]): fin.append(m)
    if fin:
        tf=doc.add_table(rows=1+len(fin),cols=1+len(periods)); tf.style='Table Grid'
        lw=2500; dw=(9360-lw)//len(periods)
        ws=[lw]+[dw]*len(periods); ws[-1]+=9360-sum(ws); _cw(tf,ws)
        for ci,h in enumerate(['Particulars']+list(periods)):
            cell=tf.rows[0].cells[ci]; _bg(cell,'1A1A2E'); _borders(cell)
            p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
            _r(p,h,bold=True,size=9,color=WHITE)
        for ri,m in enumerate(fin):
            row=tf.rows[ri+1]; bgc='FFFFFF' if ri%2==0 else 'EFF6FF'
            for ci,v in enumerate([m.get('metric','')]+[str(x) if x else '—' for x in m.get('values',[])]):
                cell=row.cells[ci]; _bg(cell,bgc); _borders(cell)
                p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
                _r(p,v,bold=(ci==0),size=9.5,color=DARK_TEXT)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Comments
    _hd(doc,'Comments')
    for c in data.get('comments',[]):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(4)
        _r(p,c['heading']+': ',bold=True,size=10.5,color=BLUE_ACCENT)
        _r(p,c['text'],size=10.5,color=DARK_TEXT)

    # Recommendation
    doc.add_paragraph().paragraph_format.space_after=Pt(4)
    _hd(doc,'Recommendation')
    p4=doc.add_paragraph(); p4.paragraph_format.space_after=Pt(16)
    _r(p4,data.get('recommendation',''),size=10.5,color=DARK_TEXT)

    # Footer
    t2=doc.add_table(rows=1,cols=2); t2.style='Table Grid'; _cw(t2,[4680,4680])
    for ci,(role,name) in enumerate([('Fund Manager',data['preparedBy']),('CIO',data['reviewedBy'])]):
        cell=t2.rows[0].cells[ci]; _bg(cell,'F1F5F9'); _borders(cell)
        p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(4)
        _r(p,role+': ',bold=True,size=10,color=MUTED_C); _r(p,name,size=10,color=DARK_TEXT)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("# 📋 Credit Report\nGenerator")
    st.markdown("---")
    st.markdown('<p class="muted">Health Insurance Premium Corpus</p>', unsafe_allow_html=True)

    st.markdown("### Configuration")
    api_key = st.text_input("Gemini API Key", type="password",
                             placeholder="AIza...",
                             help="Get free key at aistudio.google.com")

    st.markdown("---")
    st.markdown("### Entity Type")
    entity_type = st.radio("", ["Bank / SFB", "NBFC / IFC", "PSU Finance", "Apex Institution"],
                           label_visibility="collapsed")

    entity_map = {
        "Bank / SFB": "bank",
        "NBFC / IFC": "nbfc",
        "PSU Finance": "psu",
        "Apex Institution": "apex"
    }
    entity_code = entity_map[entity_type]

    st.markdown("---")
    st.markdown('<p class="muted">Cost per report: ~₹0</p>', unsafe_allow_html=True)
    st.markdown('<p class="muted">Free tier: 50 reports/day</p>', unsafe_allow_html=True)

# ── Main layout ───────────────────────────────────────────────
st.markdown("# Credit Report Generator")
st.markdown('<p class="muted" style="margin-top:-12px;margin-bottom:32px">AI-powered credit analysis for debt investment decisions</p>', unsafe_allow_html=True)

col1, col2 = st.columns([1.2, 1], gap="large")

with col1:
    st.markdown("### Issuer Details")

    c1, c2 = st.columns(2)
    with c1:
        issuer_name = st.text_input("Issuer Name", placeholder="Power Finance Corporation Ltd")
    with c2:
        review_period = st.text_input("Review Period", value="Year period ended September 30, 2025")

    c3, c4 = st.columns(2)
    with c3:
        prepared_by = st.text_input("Prepared By", placeholder="Analyst Name")
    with c4:
        reviewed_by = st.text_input("Reviewed By", placeholder="CIO Name")

    c5, c6 = st.columns(2)
    with c5:
        industry = st.text_input("Industry", value="Financial and insurance activities")
    with c6:
        sector = st.text_input("Sector", value="Finance")

    st.markdown("---")
    st.markdown("### Source Document")

    source_url = st.text_input("Document URL", placeholder="Paste direct PDF link from NSE/BSE/company IR page")

    uploaded_file = st.file_uploader("Or upload PDF", type=['pdf'],
                                      help="Upload annual report, quarterly results, or investor presentation")

    st.markdown("---")
    st.markdown("### Investments under Review")
    st.markdown('<p class="muted">Add each bond separately</p>', unsafe_allow_html=True)

    if 'investments' not in st.session_state:
        st.session_state.investments = [
            {'security':'','yield':'','agency':'','rating':'','businessHouse':'','fv':'','sh':'','ph':''}
        ]

    investments_data = []
    for i, inv in enumerate(st.session_state.investments):
        with st.expander(f"Bond {i+1}" + (f" — {inv['security']}" if inv['security'] else ""), expanded=(i==0)):
            r1c1, r1c2 = st.columns([2,1])
            with r1c1:
                sec = st.text_input("Security / Maturity", value=inv['security'],
                                     placeholder="e.g. 8.03% NCD May 02, 2026", key=f"sec_{i}")
            with r1c2:
                yld = st.text_input("Holding Yield %", value=inv['yield'],
                                     placeholder="7.55", key=f"yld_{i}")

            r2c1, r2c2, r2c3 = st.columns(3)
            with r2c1:
                agency = st.text_input("Rating Agency", value=inv['agency'],
                                        placeholder="CRISIL", key=f"agency_{i}")
            with r2c2:
                rating = st.text_input("Rating", value=inv['rating'],
                                        placeholder="AAA", key=f"rating_{i}")
            with r2c3:
                bh = st.text_input("Business House", value=inv['businessHouse'],
                                    placeholder="PFC Group", key=f"bh_{i}")

            r3c1, r3c2, r3c3 = st.columns(3)
            with r3c1:
                fv = st.text_input("Face Value (Cr)", value=inv['fv'],
                                    placeholder="25", key=f"fv_{i}")
            with r3c2:
                sh = st.text_input("SH", value=inv['sh'], placeholder="—", key=f"sh_{i}")
            with r3c3:
                ph = st.text_input("PH", value=inv['ph'], placeholder="25", key=f"ph_{i}")

            investments_data.append({
                'security':sec,'yield':yld,'agency':agency,'rating':rating,
                'businessHouse':bh,'fv':fv,'sh':sh,'ph':ph
            })

    if st.button("＋ Add another bond"):
        st.session_state.investments.append(
            {'security':'','yield':'','agency':'','rating':'','businessHouse':'','fv':'','sh':'','ph':''}
        )
        st.rerun()

with col2:
    st.markdown("### Generate Report")

    # Stats
    s1, s2, s3 = st.columns(3)
    with s1:
        st.markdown('<div class="stat-card"><div class="stat-value">~60s</div><div class="stat-label">Per Report</div></div>', unsafe_allow_html=True)
    with s2:
        st.markdown('<div class="stat-card"><div class="stat-value">₹0</div><div class="stat-label">Cost</div></div>', unsafe_allow_html=True)
    with s3:
        st.markdown(f'<div class="stat-card"><div class="stat-value">{entity_type.split("/")[0].strip()}</div><div class="stat-label">Entity Type</div></div>', unsafe_allow_html=True)

    st.markdown("")

    generate = st.button("✦ Generate Credit Report", use_container_width=True)

    if generate:
        # Validation
        errors = []
        if not api_key: errors.append("Gemini API key is required")
        if not issuer_name: errors.append("Issuer name is required")
        if not prepared_by: errors.append("Prepared By is required")
        if not reviewed_by: errors.append("Reviewed By is required")

        if errors:
            for e in errors:
                st.error(f"⚠️ {e}")
        else:
            # Extract text from sources
            extracted_text = ''

            if source_url.strip():
                with st.spinner('Fetching document from URL...'):
                    try:
                        headers = {'User-Agent': 'Mozilla/5.0'}
                        r = requests.get(source_url.strip(), headers=headers, timeout=30)
                        content_type = r.headers.get('Content-Type','').lower()
                        if 'pdf' in content_type or source_url.lower().endswith('.pdf'):
                            try:
                                import pdfplumber
                                with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                                    pages = [p.extract_text() or '' for p in pdf.pages[:40]]
                                    extracted_text = '\n'.join(pages)[:6000]
                                st.success(f"✅ PDF fetched ({len(extracted_text):,} chars)")
                            except:
                                st.warning("Could not extract PDF text from URL")
                        else:
                            from bs4 import BeautifulSoup
                            soup = BeautifulSoup(r.text, 'html.parser')
                            for tag in soup(['script','style','nav','footer']): tag.decompose()
                            extracted_text = soup.get_text(separator='\n', strip=True)[:6000]
                            st.success(f"✅ Webpage fetched ({len(extracted_text):,} chars)")
                    except Exception as e:
                        st.warning(f"Could not fetch URL: {e}")

            if uploaded_file:
                with st.spinner('Reading uploaded document...'):
                    try:
                        import pdfplumber
                        with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                            pages = [p.extract_text() or '' for p in pdf.pages[:40]]
                            file_text = '\n'.join(pages)[:6000]
                            extracted_text += f'\n{file_text}'
                        st.success(f"✅ Document read ({len(file_text):,} chars)")
                    except Exception as e:
                        st.warning(f"Could not read file: {e}")

            if not extracted_text:
                extracted_text = 'No documents provided. Use your knowledge of this company.'

            # Entity rules
            entity_rules = {
                'bank': 'Bank/SFB. INCLUDE: NII, Deposits, CASA%, GNPA/Gross Stage 3%, NNPA/Net Stage 3%, CAR%, Tier I CAR%, ROA%, ROE%, NIM%. NO Debt/Equity.',
                'nbfc': 'NBFC/IFC. INCLUDE: NII, Loans & Advances, Borrowings, Net Worth, GNPA/Gross Stage 3%, NNPA/Net Stage 3%, CRAR%, Debt/Equity, EPS, NIM%. NO Deposits.',
                'psu' : 'PSU Finance. INCLUDE: NII, PAT, Net Worth, Loans, Borrowings, GNPA/Gross Stage 3%, NNPA/Net Stage 3%, CRAR%, Debt/Equity, EPS. NO Deposits.',
                'apex': 'Apex DFI. INCLUDE: NII, PAT, Net Worth, Advances, Borrowings, GNPA%, NNPA%, CRAR%, ROA%. Comment on GoI mandate.'
            }

            inv_text = '\n'.join(
                f"- {i.get('security')} | Yield {i.get('yield')}% | {i.get('rating')} ({i.get('agency')}) | FV Rs {i.get('fv')} Cr"
                for i in investments_data if i.get('security')
            ) or 'No investments provided.'

            prompt = f"""You are a senior credit analyst at an Indian Health Insurance company (IRDAI-regulated).
Prepare a Credit Review & Analysis report.

ISSUER: {issuer_name}
ENTITY TYPE: {entity_code.upper()}
ENTITY RULES: {entity_rules[entity_code]}
INVESTMENTS: {inv_text}

SOURCE DATA:
{extracted_text}

INSTRUCTIONS:
1. Company Profile: 4-6 sentences, formal third-person, mandate/ownership/business model
2. Extract financials from source data. Calculate derived ratios:
   - GNPA/Gross Stage 3% = Gross NPA or Stage 3 Assets / Gross Advances x 100
   - NNPA/Net Stage 3% = Net NPA or Net Stage 3 / Net Advances x 100
   - NIM% = NII / Average Interest Earning Assets x 100
   - ROA% = PAT / Average Total Assets x 100
   - ROE% = PAT / Average Net Worth x 100
   - Debt/Equity = Total Borrowings / Net Worth
3. Provide all 4 periods: H1FY26, H1FY25, 31.03.2025, 31.03.2024
4. NEVER output N/A — use best estimate if not available
5. Comments: specific numbers, formal analyst tone
6. Recommendation starts with "Keeping in view..."

Respond ONLY with valid JSON, no markdown:

{{"companyProfile":"...","financialPeriods":["H1FY26","H1FY25","31.03.2025","31.03.2024"],"financialData":[{{"metric":"Total Income","values":["...","...","...","..."]}},{{"metric":"Net Interest Income","values":["...","...","...","..."]}},{{"metric":"Profit After Tax","values":["...","...","...","..."]}},{{"metric":"Net Worth","values":["...","...","...","..."]}},{{"metric":"Total Assets","values":["...","...","...","..."]}},{{"metric":"Loans & Advances","values":["...","...","...","..."]}},{{"metric":"Borrowings","values":["...","...","...","..."]}},{{"metric":"GNPA / Gross Stage 3 (%)","values":["...","...","...","..."]}},{{"metric":"NNPA / Net Stage 3 (%)","values":["...","...","...","..."]}},{{"metric":"CRAR / CAR (%)","values":["...","...","...","..."]}}],"conditionalMetrics":[{{"metric":"Deposits","values":["...","...","...","..."],"includeFor":["bank","apex"]}},{{"metric":"CASA (%)","values":["...","...","...","..."],"includeFor":["bank"]}},{{"metric":"NIM (%)","values":["...","...","...","..."],"includeFor":["bank","nbfc","psu"]}},{{"metric":"ROA (%)","values":["...","...","...","..."],"includeFor":["bank","apex"]}},{{"metric":"ROE (%)","values":["...","...","...","..."],"includeFor":["bank","nbfc","psu"]}},{{"metric":"Tier I CAR (%)","values":["...","...","...","..."],"includeFor":["bank"]}},{{"metric":"Debt/Equity","values":["...","...","...","..."],"includeFor":["nbfc","psu"]}},{{"metric":"EPS (Rs)","values":["...","...","...","..."],"includeFor":["nbfc","psu"]}}],"comments":[{{"heading":"Profitability","text":"..."}},{{"heading":"Asset Quality","text":"..."}},{{"heading":"Capitalisation","text":"..."}},{{"heading":"Liquidity","text":"..."}}],"recommendation":"Keeping in view..."}}"""

            with st.spinner('🤖 Gemini is writing your credit report...'):
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(
                        model_name='gemini-2.5-flash',
                        generation_config=genai.GenerationConfig(
                            temperature=0.1,
                            max_output_tokens=16000
                        )
                    )
                    response = model.generate_content(prompt)
                    raw = response.text

                    # Parse JSON
                    clean = re.sub(r'```json|```','',raw).strip()
                    match = re.search(r'\{[\s\S]*\}', clean)
                    if match: clean = match.group(0)
                    clean = re.sub(r',\s*}', '}', clean)
                    clean = re.sub(r',\s*]', ']', clean)

                    report_data = json.loads(clean)
                    report_data.update({
                        'issuerName':issuer_name,'entityType':entity_code,
                        'preparedBy':prepared_by,'reviewedBy':reviewed_by,
                        'industry':industry,'sector':sector,'reviewPeriod':review_period,
                        'investments':[i for i in investments_data if i.get('security')]
                    })

                    # Build DOCX
                    docx_buf = build_docx(report_data)
                    filename = issuer_name.replace(' ','_') + '_Credit_Report.docx'

                    st.success("✅ Report generated successfully!")

                    st.download_button(
                        label="⬇ Download Word Report",
                        data=docx_buf,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

                    # Preview
                    st.markdown("---")
                    st.markdown("### Preview")

                    st.markdown(f"""
                    <div class="report-card">
                        <div style="margin-bottom:8px"><span class="tag">{entity_type}</span></div>
                        <h3 style="margin:8px 0 4px">{issuer_name}</h3>
                        <p class="muted">{review_period}</p>
                        <hr>
                        <p style="font-size:0.9rem;color:#c8c4bc;line-height:1.7">{report_data.get('companyProfile','')}</p>
                        <hr>
                        <p style="font-size:0.85rem;color:#9b9688"><strong style="color:#c9a84c">Recommendation:</strong> {report_data.get('recommendation','')[:200]}...</p>
                    </div>
                    """, unsafe_allow_html=True)

                except json.JSONDecodeError:
                    st.error("❌ JSON parse error — click Generate again")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
