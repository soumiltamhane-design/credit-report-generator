# ============================================================
# APP.PY — UPGRADED MULTIMODAL CREDIT REPORT GENERATOR WEB UI
# ============================================================
import streamlit as st
import json
import re
import os
import time
import io
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── STREAMLIT PAGE CONFIG ───────────────────────────────────
st.set_page_config(
    page_title="Credit Review Automation Engine",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Financial-Themed CSS Shading (Dark Navy & Blue Accent)
st.markdown("""
    <style>
    .reportview-container { background: #F8FAFC; }
    .sidebar .sidebar-content { background: #1A1A2E; }
    h1, h2, h3 { color: #1A1A2E; font-family: 'Calibri', sans-serif; }
    div.stButton > button:first-child {
        background-color: #2563EB;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        border: none;
        padding: 0.5rem 2rem;
    }
    div.stButton > button:first-child:hover { background-color: #1D4ED8; }
    </style>
""", unsafe_allow_html=True)

# ── DOCX DESIGN TEMPLATE HELPERS ────────────────────────────
def _bg(cell, h):
    s=OxmlElement('w:shd'); s.set(qn('w:val'),'clear')
    s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),h)
    cell._tc.get_or_add_tcPr().append(s)

def _borders(cell):
    b=OxmlElement('w:tcBorders')
    for side in ('top','bottom','left','right'):
        e=OxmlElement(f'w:{side}'); e.set(qn('w:val'),'single')
        e.set(qn('w:sz'),'2'); e.set(qn('w:space'),'0')
        e.set(qn('w:color'),'CCCCCC'); b.append(e)
    cell._tc.get_or_add_tcPr().append(b)

def _cw(table, widths):
    tg=OxmlElement('w:tblGrid')
    for w in widths:
        gc=OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(w)); tg.append(gc)
    table._tbl.insert(1,tg)
    for row in table.rows:
        for cell,w in zip(row.cells,widths):
            tw=OxmlElement('w:tcW'); tw.set(qn('w:w'),str(w))
            tw.set(qn('w:type'),'dxa'); cell._tc.get_or_add_tcPr().append(tw)

def _r(para, text, bold=False, size=10.5, color=None):
    run=para.add_run(str(text)); run.bold=bold
    run.font.name='Calibri'; run.font.size=Pt(size)
    if color: run.font.color.rgb=color
    return run

def _hd(doc, text):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(12)
    p.paragraph_format.space_after=Pt(4)
    _r(p,text.upper(),bold=True,size=9.5,color=RGBColor(0x25,0x63,0xEB))
    pBdr=OxmlElement('w:pBdr'); btm=OxmlElement('w:bottom')
    btm.set(qn('w:val'),'single'); btm.set(qn('w:sz'),'8')
    btm.set(qn('w:space'),'1'); btm.set(qn('w:color'),'2563EB')
    pBdr.append(btm); p._p.get_or_add_pPr().append(pBdr)

def build_docx(data):
    doc=Document()
    for section in doc.sections:
        section.page_height=Cm(29.7); section.page_width=Cm(21.0)
        for margin in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin'):
            setattr(section, margin, Inches(1))

    entity=data['entityType']
    periods=data.get('financialPeriods',['H1FY26','H1FY25','31.03.2025','31.03.2024'])

    # Title Block
    for text,sz,fill,clr in [
        (data['issuerName'].upper(),18,'1A1A2E',RGBColor(255,255,255)),
        ('Credit Review & Analysis',12,'1A1A2E',RGBColor(0x94,0xA3,0xB8)),
        (' ',4,'2563EB',RGBColor(255,255,255))
    ]:
        p=doc.add_paragraph()
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
        shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear')
        shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),fill)
        p._p.get_or_add_pPr().append(shd)
        _r(p,text,bold=(sz==18),size=sz,color=clr)

    _hd(doc,'Company Profile')
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(8)
    _r(p,data.get('companyProfile',''),size=10.5,color=RGBColor(0x1F,0x29,0x37))

    _hd(doc,'Issuer Information')
    t=doc.add_table(rows=6,cols=2); t.style='Table Grid'; _cw(t,[2700,6300])
    info_pairs = [
        ('Issuer', data['issuerName']), ('Industry as per NIC code', data.get('industry','')),
        ('Sector', data.get('sector','')), ('Review Period', data.get('reviewPeriod','')),
        ('Prepared By', data['preparedBy']), ('Reviewed By', data['reviewedBy'])
    ]
    for i,(label,value) in enumerate(info_pairs):
        lc,vc=t.rows[i].cells[0],t.rows[i].cells[1]
        _bg(lc,'F1F5F9'); _borders(lc); _bg(vc,'FFFFFF'); _borders(vc)
        for cell,txt,bold,clr in [(lc,label,True,RGBColor(0x47,0x55,0x69)),(vc,value,False,RGBColor(0x1F,0x29,0x37))]:
            p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
            _r(p,txt,bold=bold,size=9.5,color=clr)

    _hd(doc,'Brief Detail of Investment under Review')
    p2=doc.add_paragraph(); p2.paragraph_format.space_after=Pt(4)
    _r(p2,'(Rs. in crores)',size=9,color=RGBColor(0x47,0x55,0x69))
    invs=[i for i in data.get('investments',[]) if i.get('security')]
    if invs:
        hdrs=['Security / Date of Maturity','Holding Yield','Credit Rating Agency','Credit Rating','Business House','Total Nominal Value (F.V)','SH','PH']
        ti=doc.add_table(rows=1+len(invs),cols=8); ti.style='Table Grid'
        ws=[2100,680,1100,700,1100,1100,680,680]
        total=9360; s=sum(ws); ws=[int(w*total/s) for w in ws]; ws[-1]+=total-sum(ws)
        _cw(ti,ws)
        for ci,h in enumerate(hdrs):
            cell=ti.rows[0].cells[ci]; _bg(cell,'1A1A2E'); _borders(cell)
            p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
            _r(p,h,bold=True,size=8,color=RGBColor(255,255,255))
        for ri,inv in enumerate(invs):
            row=ti.rows[ri+1]; bgc='FFFFFF' if ri%2==0 else 'F8FAFC'
            vals = [inv.get('security',''), inv.get('yield',''), inv.get('agency',''), inv.get('rating',''), inv.get('businessHouse',''), inv.get('fv',''), inv.get('sh',''), inv.get('ph','')]
            for ci,v in enumerate(vals):
                cell=row.cells[ci]; _bg(cell,bgc); _borders(cell)
                p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
                _r(p,v or '',size=9,color=RGBColor(0x1F,0x29,0x37))

    _hd(doc,'Financial Strength')
    p3=doc.add_paragraph(); p3.paragraph_format.space_after=Pt(4)
    _r(p3,'(Rs. in crores unless stated otherwise)',size=9,color=RGBColor(0x47,0x55,0x69))
    fin=list(data.get('financialData',[]))
    for m in data.get('conditionalMetrics',[]):
        if entity in m.get('includeFor',[]): fin.append(m)
    if fin:
        tf=doc.add_table(rows=1+len(fin),cols=1+len(periods)); tf.style='Table Grid'
        lw=2500; dw=(9360-lw)//len(periods); ws=[lw]+[dw]*len(periods); ws[-1]+=9360-sum(ws); _cw(tf,ws)
        
        header_row = ['Particulars'] + list(periods)
        for ci,h in enumerate(header_row):
            if ci < len(tf.rows[0].cells):
                cell=tf.rows[0].cells[ci]; _bg(cell,'1A1A2E'); _borders(cell)
                p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
                _r(p,h,bold=True,size=9,color=RGBColor(255,255,255))
                
        for ri,m in enumerate(fin):
            row=tf.rows[ri+1]; bgc='FFFFFF' if ri%2==0 else 'EFF6FF'
            
            # Boundary Protection Check: Safeguards structure against response dynamic length variances
            raw_vals = m.get('values', [])
            sanitized_vals = []
            for idx in range(len(periods)):
                if idx < len(raw_vals):
                    val = str(raw_vals[idx])
                    sanitized_vals.append(val if val not in [None, '', 'nan'] else '—')
                else:
                    sanitized_vals.append('—')
            
            m_row = [m.get('metric','')] + sanitized_vals
            for ci,v in enumerate(m_row):
                if ci < len(row.cells):
                    cell=row.cells[ci]; _bg(cell,bgc); _borders(cell)
                    p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
                    _r(p,v,bold=(ci==0),size=9.5,color=RGBColor(0x1F,0x29,0x37))

    _hd(doc,'Comments')
    for c in data.get('comments',[]):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(4)
        _r(p,c.get('heading','')+': ',bold=True,size=10.5,color=RGBColor(0x25,0x63,0xEB))
        _r(p,c.get('text',''),size=10.5,color=RGBColor(0x1F,0x29,0x37))

    _hd(doc,'Recommendation')
    p4=doc.add_paragraph(); p4.paragraph_format.space_after=Pt(16)
    _r(p4,data.get('recommendation',''),size=10.5,color=RGBColor(0x1F,0x29,0x37))

    t2=doc.add_table(rows=1,cols=2); t2.style='Table Grid'; _cw(t2,[4680,4680])
    for ci,(role,name) in enumerate([('Fund Manager',data['preparedBy']),('CIO',data['reviewedBy'])]):
        cell=t2.rows[0].cells[ci]; _bg(cell,'F1F5F9'); _borders(cell)
        p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(4)
        _r(p,role+': ',bold=True,size=10,color=RGBColor(0x47,0x55,0x69)); _r(p,name,size=10,color=RGBColor(0x1F,0x29,0x37))
    return doc

# ── STREAMLIT FRONTEND LAYOUT ───────────────────────────────
st.title("📋 Automated Credit Review & Analysis Platform")
st.subheader("Health Insurance Premium Corpus — Institutional Debt Automation")

# SIDEBAR: SETUP & CREDENTIALS
st.sidebar.header("🔑 Authentication & Sector Setup")
api_key_input = st.sidebar.text_input("Gemini API Key", type="password", value=os.environ.get("GEMINI_API_KEY", ""))
entity_type = st.sidebar.selectbox("Target Entity Classification", ["nbfc", "bank", "psu", "apex"], index=0)

# MAIN FORM: METADATA
st.subheader("1. General Assignment Parameters")
c1, c2, c3 = st.columns(3)
issuer_name = c1.text_input("Issuer Corporate Name", "Aditya Birla Capital Ltd")
prepared_by = c2.text_input("Prepared By (Analyst / Fund Manager)", "Harshal Shah")
reviewed_by = c3.text_input("Reviewed By (CIO)", "Dhaval Shah")

c4, c5, c6 = st.columns(3)
industry = c4.text_input("Industry Classification Code", "NBFC - Core Investment Company")
sector = c5.text_input("Investment Sector", "Financial Services")
review_period = c6.text_input("Report Analysis Review Period", "H1FY26 Review")

# MAIN FORM: INTERACTIVE DATA EDITOR FOR BONDS
st.subheader("2. Details of Debt Security / Investment under Review")
default_bonds = [
    {"security": "NCD 04.02.2026", "yield": "7.54", "agency": "CRISIL", "rating": "A1+", "businessHouse": "Axis Bank", "fv": "25", "sh": "", "ph": "25"},
    {"security": "NCD 08.01.2026", "yield": "7.45", "agency": "CRISIL", "rating": "A1+", "businessHouse": "Axis Bank", "fv": "25", "sh": "", "ph": "25"}
]
bond_df = pd.DataFrame(default_bonds)
edited_bond_df = st.data_editor(bond_df, num_rows="dynamic", use_container_width=True)

# MAIN FORM: UPGRADED MULTI-FILE UPLOADER
st.subheader("3. Curated Summary Evidence Source Documents")
uploaded_files = st.file_uploader(
    "Select and drop your financial sheets / summary PDFs (You can select multiple files at once)", 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

# ── COMPILING ENGINE EXECUTION ───────────────────────────────
if st.button("Generate Formal Credit Document"):
    if not api_key_input.strip():
        st.error("❌ Action Blocked: Please enter a valid Gemini API Key in the sidebar.")
    elif not uploaded_files:
        st.error("❌ Action Blocked: Please upload at least one curated document page or summary sheet.")
    else:
        with st.spinner("🚀 Initializing native File API hosting engine..."):
            genai.configure(api_key=api_key_input.strip())
            hosted_gemini_files = []
            
            try:
                # Host each uploaded memory file binary directly to Gemini server layout channels
                for idx, file in enumerate(uploaded_files, start=1):
                    ext = file.name.split('.')[-1].lower()
                    mime_type = 'application/pdf' if ext == 'pdf' else f'image/{ext}'
                    
                    tmp_filepath = f"/tmp/web_upload_{idx}.{ext}"
                    with open(tmp_filepath, "wb") as f:
                        f.write(file.getbuffer())
                    
                    st.write(f"📤 Hosting item: `{file.name}` directly on Gemini server cluster...")
                    gfile = genai.upload_file(tmp_filepath, mime_type=mime_type)
                    hosted_gemini_files.append(gfile)
                
                st.success(f"✅ Context packages active on server nodes: {len(hosted_gemini_files)} files hosted.")
                
            except Exception as e:
                st.error(f"❌ Connection Interrupted during server file injection: {e}")
                st.stop()

        with st.spinner("🤖 Executing credit reasoning logic model evaluation (Est. 20-30 seconds)..."):
            ENTITY_RULES = {
                'bank': 'Bank/SFB. INCLUDE: NII, Deposits, CASA%, GNPA/Gross Stage 3%, NNPA/Net Stage 3%, CAR%, Tier I CAR%, ROA%, ROE%, NIM%. NO Debt/Equity.',
                'nbfc': 'NBFC/IFC. INCLUDE: NII, Loans & Advances, Borrowings, Net Worth, GNPA/Gross Stage 3%, NNPA/Net Stage 3%, CRAR%, Debt/Equity, EPS, NIM%. NO Deposits.',
                'psu' : 'PSU Finance. INCLUDE: NII, PAT, Net Worth, Loans, Borrowings, GNPA/Gross Stage 3%, NNPA/Net Stage 3%, CRAR%, Debt/Equity, EPS. NO Deposits.',
                'apex': 'Apex DFI. INCLUDE: NII, PAT, Net Worth, Advances, Borrowings, GNPA%, NNPA%, CRAR%, ROA%. Comment on GoI mandate.'
            }
            
            investments_list = edited_bond_df.to_dict(orient="records")
            inv_text = "\n".join([f"- {i.get('security')} | Yield {i.get('yield')}% | {i.get('rating')} ({i.get('agency')}) | FV Rs {i.get('fv')} Cr" for i in investments_list if i.get('security')])
            
            # Formulate the Strict Alignment Column Anchor Prompt
            PROMPT = f"""You are a senior credit analyst at an Indian Health Insurance company.
Prepare a Credit Review & Analysis report based on the attached raw financial PDF elements.

ISSUER: {issuer_name}
ENTITY TYPE: {entity_type.upper()}
ENTITY RULES: {ENTITY_RULES[entity_type]}
INVESTMENTS: {inv_text}

CRITICAL EXTRACTION INSTRUCTIONS (STRICT GROUNDING):
1) NATIVE PDF LAYOUT MAPPING:
   - The attached document contains data arranged side-by-side in vertical columns. 
   - Verify column headings visually. Map figures carefully across the 4 JSON data indices:
     * Index 0 = H1FY26 ("Sep 2025", "H1FY26", or "6 Months Ended Sep 30, 2025")
     * Index 1 = H1FY25 ("Sep 2024", "H1FY25", or "6 Months Ended Sep 30, 2024")
     * Index 2 = 31.03.2025 ("Mar 2025", "FY25", or "Year Ended March 31, 2025")
     * Index 3 = 31.03.2024 ("Mar 2024", "FY24", or "Year Ended March 31, 2024")
   - IF A PERIOD COLUMN IS MISSING IN THE TARGET FILE, OUTPUT "—" FOR THAT INDEX. Never shift adjacent columns.

2) ROW ANCHORING:
   - Extract metrics exactly as listed. Do not look at row titles that don't match your target parameters.
   - Execute clean unit conversions (e.g., if sheet presents values in Lakhs, divide by 100 to convert to Crores).

3) NARRATIVE SECTIONS:
   - Company Profile: 4-6 formal third-person sentences overview.
   - Comments: write technical paragraphs for Profitability, Asset Quality, Capitalisation, and Liquidity referencing exact numbers.
   - Recommendation: start with "Keeping in view..." and maintain an objective credit perspective.

Respond ONLY with valid JSON. No markdown. No explanation. Just the JSON object.

{{
  "companyProfile": "...",
  "financialPeriods": ["H1FY26","H1FY25","31.03.2025","31.03.2024"],
  "financialData": [
    {{"metric":"Total Income","values":["...","...","...","..."]}},
    {{"metric":"Net Interest Income","values":["...","...","...","..."]}},
    {{"metric":"Profit After Tax","values":["...","...","...","..."]}},
    {{"metric":"Net Worth","values":["...","...","...","..."]}},
    {{"metric":"Total Assets","values":["...","...","...","..."]}},
    {{"metric":"Loans & Advances","values":["...","...","...","..."]}},
    {{"metric":"Borrowings","values":["...","...","...","..."]}},
    {{"metric":"GNPA / Gross Stage 3 (%)","values":["...","...","...","..."]}},
    {{"metric":"NNPA / Net Stage 3 (%)","values":["...","...","...","..."]}},
    {{"metric":"CRAR / CAR (%)","values":["...","...","...","..."]}}
  ],
  "conditionalMetrics": [
    {{"metric":"Deposits","values":["...","...","...","..."],"includeFor":["bank","apex"]}},
    {{"metric":"CASA (%)","values":["...","...","...","..."],"includeFor":["bank"]}},
    {{"metric":"NIM (%)","values":["...","...","...","..."],"includeFor":["bank","nbfc","psu"]}},
    {{"metric":"ROA (%)","values":["...","...","...","..."],"includeFor":["bank","apex"]}},
    {{"metric":"ROE (%)","values":["...","...","...","..."],"includeFor":["bank","nbfc","psu"]}},
    {{"metric":"Tier I CAR (%)","values":["...","...","...","..."],"includeFor":["bank"]}},
    {{"metric":"Debt/Equity","values":["...","...","...","..."],"includeFor":["nbfc","psu"]}},
    {{"metric":"EPS (Rs)","values":["...","...","...","..."],"includeFor":["nbfc","psu"]}}
  ],
  "comments": [
    {{"heading":"Profitability","text":"..."}},
    {{"heading":"Asset Quality","text":"..."}},
    {{"heading":"Capitalisation","text":"..."}},
    {{"heading":"Liquidity","text":"..."}}
  ],
  "recommendation": "Keeping in view...",
  "dataQualityWarnings": []
}}"""

            # Build content request structure
            content_payload = [PROMPT] + hosted_gemini_files
            model = genai.GenerativeModel(
                model_name='gemini-2.5-flash',
                generation_config=genai.GenerationConfig(temperature=0.0, response_mime_type="application/json")
            )
            
            raw_response_text = None
            for attempt in range(3):
                try:
                    response = model.generate_content(content_payload)
                    raw_response_text = response.text
                    break
                except Exception as e:
                    if "429" in str(e) or "quota" in str(e).lower():
                        st.warning(f"⚠️ Traffic restriction encountered. Automatically sleeping 60 seconds to clear window (Attempt {attempt+1}/3)...")
                        time.sleep(60)
                    else:
                        st.error(f"❌ Model operational error: {e}")
                        st.stop()
            
            if not raw_response_text:
                st.error("❌ Transaction Timed Out: Traffic capacity limits reached on free hosting tier nodes.")
                st.stop()

        with st.spinner("📝 Assembling structural layout to styled Word document formatting layer..."):
            try:
                # Clean structural markdown
                clean_json_str = re.sub(r'```json|```', '', raw_response_text).strip()
                match = re.search(r'\{[\s\S]*\}', clean_json_str)
                if match: clean_json_str = match.group(0)
                
                report_data = json.loads(clean_json_str)
                
                # Combine system parameters
                report_data.update({
                    'issuerName': issuer_name, 'entityType': entity_type,
                    'preparedBy': prepared_by, 'reviewedBy': reviewed_by,
                    'industry': industry, 'sector': sector, 'reviewPeriod': review_period,
                    'investments': investments_list
                })
                
                # Construct template binary
                generated_docx = build_docx(report_data)
                
                docx_buffer = io.BytesIO()
                generated_docx.save(docx_buffer)
                docx_bytes = docx_buffer.getvalue()
                
                st.success("🎉 Credit Review compiled completely without a single verification crash!")
                
                # Render UI download portal interface buttons
                filename_output = f"{issuer_name.replace(' ', '_')}_Credit_Report.docx"
                st.download_button(
                    label="📥 Download Formatted Credit Report (.docx)",
                    data=docx_bytes,
                    file_name=filename_output,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"❌ Synthesis Interrupted during document structuring: {e}")
                st.text_area("System Trace Diagnostic Data Payload", value=raw_response_text, height=250)
