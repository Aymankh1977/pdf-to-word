import io
import re
import os
import zipfile
import streamlit as st
import pdfplumber
import pikepdf
from pypdf import PdfReader, PdfWriter
from pdf2image import convert_from_bytes
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import Color, HexColor

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF Studio", page_icon="📄", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,600;1,300;1,400&family=Jost:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'Jost', sans-serif; letter-spacing: 0.01em; }

.stApp { background: #f7f3ee; color: #1c1917; }

[data-testid="stSidebar"] { background: #1c1917 !important; border-right: none !important; }
[data-testid="stSidebar"] * { color: #e8e0d5 !important; }
[data-testid="stSidebar"] .stRadio label {
    font-family: 'Jost', sans-serif !important; font-size: 0.8rem !important;
    font-weight: 400 !important; letter-spacing: 0.1em !important;
    text-transform: uppercase !important; padding: 0.45rem 0 !important;
    color: #7a6e64 !important; border-bottom: 1px solid #2a2420 !important;
    display: block; width: 100%; transition: color 0.2s;
}
[data-testid="stSidebar"] .stRadio label:hover { color: #e8e0d5 !important; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: #4a4038 !important; font-size: 0.75rem !important; line-height: 1.6 !important; }
[data-testid="stSidebar"] hr { border-color: #2a2420 !important; margin: 1rem 0 !important; }

.hero { text-align: center; padding: 3.5rem 1rem 2.5rem; border-bottom: 1px solid #e2d9ce; margin-bottom: 2.5rem; }
.hero-eyebrow { font-family: 'Jost', sans-serif; font-size: 0.68rem; font-weight: 500; letter-spacing: 0.28em; text-transform: uppercase; color: #a89880; margin-bottom: 0.9rem; }
.hero h1 { font-family: 'Cormorant Garamond', serif !important; font-size: 3.8rem !important; font-weight: 300 !important; color: #1c1917 !important; letter-spacing: -0.02em !important; line-height: 1.05 !important; margin-bottom: 0.8rem !important; }
.hero h1 em { font-style: italic; color: #8b5e52; }
.hero-sub { font-size: 0.78rem; color: #9a8e82; font-weight: 300; letter-spacing: 0.08em; }

.tool-heading { display: flex; align-items: baseline; gap: 0.8rem; margin-bottom: 2rem; padding-bottom: 1.2rem; border-bottom: 1px solid #e2d9ce; }
.tool-heading .tool-icon { font-size: 1.1rem; }
.tool-heading h2 { font-family: 'Cormorant Garamond', serif !important; font-size: 2rem !important; font-weight: 400 !important; color: #1c1917 !important; margin: 0 !important; }
.tool-heading .tool-tag { font-size: 0.65rem; font-weight: 500; letter-spacing: 0.15em; text-transform: uppercase; color: #a89880; background: #ede6dc; padding: 3px 10px; border-radius: 20px; margin-left: 4px; }

.sidebar-logo { font-family: 'Cormorant Garamond', serif; font-size: 1.5rem; font-weight: 300; color: #e8e0d5; letter-spacing: 0.04em; padding: 1.5rem 0 1.2rem; border-bottom: 1px solid #2a2420; margin-bottom: 1rem; }
.sidebar-logo em { font-style: italic; color: #c4967a; }
.sidebar-section { font-size: 0.6rem; letter-spacing: 0.22em; text-transform: uppercase; color: #4a4038 !important; margin: 1.2rem 0 0.6rem; font-weight: 600; }

[data-testid="stFileUploader"] { background: #ffffff !important; border: 1px solid #d5ccc4 !important; border-radius: 2px !important; }
[data-testid="stFileUploader"]:hover { border-color: #8b5e52 !important; }

.stButton > button {
    background: #1c1917 !important; color: #f7f3ee !important; border: none !important;
    border-radius: 2px !important; font-family: 'Jost', sans-serif !important;
    font-weight: 500 !important; font-size: 0.72rem !important;
    letter-spacing: 0.14em !important; text-transform: uppercase !important;
    padding: 0.7rem 2.2rem !important; transition: background 0.2s !important;
}
.stButton > button:hover { background: #3d3530 !important; }

[data-testid="stDownloadButton"] > button {
    background: transparent !important; color: #1c1917 !important;
    border: 1.5px solid #1c1917 !important; border-radius: 2px !important;
    font-family: 'Jost', sans-serif !important; font-weight: 500 !important;
    font-size: 0.72rem !important; letter-spacing: 0.14em !important;
    text-transform: uppercase !important; padding: 0.65rem 1.5rem !important;
    width: 100% !important; transition: all 0.2s !important;
}
[data-testid="stDownloadButton"] > button:hover { background: #1c1917 !important; color: #f7f3ee !important; }

[data-testid="stInfo"] { background: #ede6dc !important; border: none !important; border-left: 3px solid #a89880 !important; border-radius: 0 !important; color: #1c1917 !important; }
[data-testid="stSuccess"] { background: #e8ede6 !important; border: none !important; border-left: 3px solid #6b8b5e !important; border-radius: 0 !important; }
[data-testid="stError"] { background: #f0e6e6 !important; border: none !important; border-left: 3px solid #8b5e5e !important; border-radius: 0 !important; }

.result-card { background: #ffffff; border: 1px solid #e2d9ce; border-left: 3px solid #8b5e52; border-radius: 0 2px 2px 0; padding: 1rem 1.4rem; margin-bottom: 0.6rem; }
.result-card .fname { font-family: 'Jost', sans-serif; font-weight: 500; font-size: 0.88rem; color: #1c1917; margin-bottom: 3px; }
.result-card .fmeta { font-size: 0.76rem; color: #9a8e82; font-weight: 300; }

.stProgress > div > div { background: #8b5e52 !important; }

.stTextInput input, .stNumberInput input {
    background: #ffffff !important; border: 1px solid #d5ccc4 !important;
    border-radius: 2px !important; color: #1c1917 !important;
    font-family: 'Jost', sans-serif !important; font-size: 0.88rem !important;
}
.stTextInput input:focus, .stNumberInput input:focus { border-color: #8b5e52 !important; box-shadow: 0 0 0 1px #8b5e52 !important; }

[data-baseweb="select"] > div { background: #ffffff !important; border: 1px solid #d5ccc4 !important; border-radius: 2px !important; }

label[data-testid="stWidgetLabel"] {
    font-family: 'Jost', sans-serif !important; font-size: 0.72rem !important;
    font-weight: 500 !important; letter-spacing: 0.12em !important;
    text-transform: uppercase !important; color: #6b5f55 !important;
}
.stCaption, [data-testid="stCaptionContainer"] { font-size: 0.73rem !important; color: #a89880 !important; font-style: italic !important; }

hr { border-color: #e2d9ce !important; margin: 1.5rem 0 !important; }
#MainMenu, footer, [data-testid="stToolbar"] { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
#  UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════════

def get_pdf_info(pdf_bytes):
    r = PdfReader(io.BytesIO(pdf_bytes))
    return {"pages": len(r.pages), "encrypted": r.is_encrypted}

def pdf_page_previews(pdf_bytes, dpi=80, max_pages=20):
    """Return list of PIL images for preview."""
    try:
        imgs = convert_from_bytes(pdf_bytes, dpi=dpi, first_page=1, last_page=max_pages)
        return imgs
    except Exception:
        return []

# ── 1. PDF → Word ─────────────────────────────────────────────────

def clean_text(text):
    if not text: return ""
    text = text.replace('\xa0', ' ')
    return re.sub(r'[ \t]{2,}', ' ', text).strip()

def looks_like_heading(text, font_size=None, is_bold=False):
    s = text.strip()
    if not s: return 0
    if s.isupper() and 3 <= len(s) <= 80: return 1
    if is_bold and len(s) <= 80 and not s.endswith('.'): return 2
    if font_size:
        if font_size >= 18: return 1
        if font_size >= 14: return 2
        if font_size >= 12 and is_bold: return 3
    return 0

def set_table_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'4')
        b.set(qn('w:space'),'0'); b.set(qn('w:color'),'AAAAAA')
        tblBorders.append(b)
    tblPr.append(tblBorders)

def add_table_to_doc(doc, table_data, font_size_pt=11):
    rows = [r for r in table_data if any(c for c in r)]
    if not rows: return
    num_cols = max(len(r) for r in rows)
    tbl = doc.add_table(rows=len(rows), cols=num_cols)
    tbl.style = 'Table Grid'
    for ri, row in enumerate(rows):
        for ci in range(num_cols):
            ct = row[ci] if ci < len(row) else ""
            cell = tbl.cell(ri, ci)
            cell.text = clean_text(ct or "")
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(font_size_pt - 1)
                if ri == 0: run.bold = True
            if ri == 0:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'E8E8E8')
                tcPr.append(shd)
    set_table_border(tbl)
    doc.add_paragraph()

def extract_page_content(page):
    content = []
    tables = page.extract_tables()
    try: table_bboxes = [t.bbox for t in page.find_tables()]
    except: table_bboxes = []
    try: image_bboxes = [(i['x0'],i['top'],i['x1'],i['bottom']) for i in page.images]
    except: image_bboxes = []
    words = page.extract_words(extra_attrs=["size","fontname"])
    if not words:
        raw = page.extract_text()
        if raw:
            for line in raw.splitlines():
                line = clean_text(line)
                if line: content.append({"type":"paragraph","text":line,"font_size":None,"bold":False})
        for tbl in tables:
            if tbl: content.append({"type":"table","data":tbl})
        return content, image_bboxes
    lines = {}
    for w in words: lines.setdefault(round(w['top'],1),[]).append(w)
    def in_r(y,bbs): return any(b[1]<=y<=b[3] for b in bbs)
    prev_bottom = None
    for top, lw in sorted(lines.items()):
        if in_r(top,table_bboxes) or in_r(top,image_bboxes): continue
        lw.sort(key=lambda w: w['x0'])
        lt = clean_text(' '.join(w['text'] for w in lw))
        if not lt: continue
        sizes = [w.get('size',0) for w in lw if w.get('size')]
        fs = max(sizes) if sizes else None
        bold = any('Bold' in w.get('fontname','') for w in lw)
        if prev_bottom and (top-prev_bottom)>12: content.append({"type":"blank"})
        content.append({"type":"paragraph","text":lt,"font_size":fs,"bold":bold})
        prev_bottom = top + lw[0].get('height',10)
    for tbl in tables:
        if tbl: content.append({"type":"table","data":tbl})
    return content, image_bboxes

def crop_img(page_img, bbox, pw, ph):
    iw, ih = page_img.size
    sx, sy = iw/pw, ih/ph
    x0,y0,x1,y1 = int(bbox[0]*sx),int(bbox[1]*sy),int(bbox[2]*sx),int(bbox[3]*sy)
    pad=4; x0,y0=max(0,x0-pad),max(0,y0-pad); x1,y1=min(iw,x1+pad),min(ih,y1+pad)
    if x1<=x0 or y1<=y0: return None
    return page_img.crop((x0,y0,x1,y1))

def convert_pdf_to_docx(pdf_bytes, font_size=11, include_images=True, ocr_fallback=True, dpi=150):
    doc = Document()
    for s in doc.sections:
        s.top_margin=s.bottom_margin=Inches(1)
        s.left_margin=s.right_margin=Inches(1.2)
    doc.styles['Normal'].font.name='Calibri'
    doc.styles['Normal'].font.size=Pt(font_size)
    page_images=[]
    if include_images:
        try: page_images=convert_from_bytes(pdf_bytes,dpi=dpi)
        except: pass
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total=len(pdf.pages)
        for pi, page in enumerate(pdf.pages):
            content, image_bboxes = extract_page_content(page)
            raw_text = page.extract_text() or ""
            if not raw_text.strip() and ocr_fallback and page_images:
                try:
                    import pytesseract
                    if pi < len(page_images):
                        for line in pytesseract.image_to_string(page_images[pi]).splitlines():
                            line=clean_text(line)
                            if line: content.append({"type":"paragraph","text":line,"font_size":None,"bold":False})
                except: pass
            for item in content:
                if item["type"]=="blank": continue
                elif item["type"]=="paragraph":
                    text=item["text"]
                    level=looks_like_heading(text,item.get("font_size"),item.get("bold",False))
                    if level in (1,2,3):
                        h=doc.add_heading(text,level=level)
                        h.runs[0].font.size=Pt(font_size+(6-level*2))
                    else:
                        p=doc.add_paragraph(text)
                        for run in p.runs:
                            run.font.size=Pt(font_size)
                            if item.get("bold"): run.bold=True
                elif item["type"]=="table":
                    add_table_to_doc(doc,item["data"],font_size)
            if include_images and page_images and pi<len(page_images):
                pg_img=page_images[pi]; used=[]
                for bbox in image_bboxes:
                    w,h=bbox[2]-bbox[0],bbox[3]-bbox[1]
                    if w<20 or h<20: continue
                    if any(not(bbox[2]<u[0] or bbox[0]>u[2] or bbox[3]<u[1] or bbox[1]>u[3]) for u in used): continue
                    used.append(bbox)
                    cropped=crop_img(pg_img,bbox,page.width,page.height)
                    if not cropped: continue
                    ib=io.BytesIO(); cropped.save(ib,format='PNG'); ib.seek(0)
                    iw2,ih2=cropped.size; mw=Inches(5)
                    asp=ih2/iw2 if iw2>0 else 1; dw=min(mw,Inches(iw2/dpi))
                    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(ib,width=dw); doc.add_paragraph()
            if pi<total-1: doc.add_page_break()
    out=io.BytesIO(); doc.save(out); return out.getvalue()

# ── 2. Merge PDFs ─────────────────────────────────────────────────

def merge_pdfs(pdf_list):
    writer = PdfWriter()
    for pdf_bytes in pdf_list:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

# ── 3. Split PDF ──────────────────────────────────────────────────

def split_pdf(pdf_bytes, ranges):
    """ranges: list of (start, end) 1-indexed inclusive tuples"""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    results = []
    for start, end in ranges:
        writer = PdfWriter()
        for i in range(start-1, min(end, len(reader.pages))):
            writer.add_page(reader.pages[i])
        out = io.BytesIO()
        writer.write(out)
        results.append((f"pages_{start}_to_{end}.pdf", out.getvalue()))
    return results

# ── 4. Rotate pages ───────────────────────────────────────────────

def rotate_pdf(pdf_bytes, angle, page_nums=None):
    """angle: 90, 180, 270. page_nums: list of 1-indexed, or None for all."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if page_nums is None or (i+1) in page_nums:
            page.rotate(angle)
        writer.add_page(page)
    out = io.BytesIO(); writer.write(out); return out.getvalue()

# ── 5. Watermark ──────────────────────────────────────────────────

def add_watermark(pdf_bytes, text, opacity=0.25, font_size=60, color_hex="#cc3333", angle=45):
    # Build watermark page
    wm_buf = io.BytesIO()
    c = rl_canvas.Canvas(wm_buf, pagesize=letter)
    r = int(color_hex[1:3],16)/255
    g = int(color_hex[3:5],16)/255
    b = int(color_hex[5:7],16)/255
    c.setFillColor(Color(r, g, b, alpha=opacity))
    c.setFont("Helvetica-Bold", font_size)
    c.saveState()
    c.translate(letter[0]/2, letter[1]/2)
    c.rotate(angle)
    c.drawCentredString(0, 0, text)
    c.restoreState()
    c.save()
    wm_page = PdfReader(io.BytesIO(wm_buf.getvalue())).pages[0]

    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for page in reader.pages:
        page.merge_page(wm_page)
        writer.add_page(page)
    out = io.BytesIO(); writer.write(out); return out.getvalue()

# ── 6. Password protect / unlock ─────────────────────────────────

def protect_pdf(pdf_bytes, user_password, owner_password=None):
    writer = PdfWriter()
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for page in reader.pages: writer.add_page(page)
    writer.encrypt(user_password, owner_password or user_password)
    out = io.BytesIO(); writer.write(out); return out.getvalue()

def unlock_pdf(pdf_bytes, password):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if reader.is_encrypted:
        reader.decrypt(password)
    writer = PdfWriter()
    for page in reader.pages: writer.add_page(page)
    out = io.BytesIO(); writer.write(out); return out.getvalue()

# ── 7. Compress ───────────────────────────────────────────────────

def compress_pdf(pdf_bytes):
    out = io.BytesIO()
    with pikepdf.open(io.BytesIO(pdf_bytes)) as pk:
        pk.save(out, compress_streams=True, recompress_flate=True)
    return out.getvalue()

# ── 8. Extract pages as images ────────────────────────────────────

def extract_as_images(pdf_bytes, dpi=150, fmt="PNG"):
    imgs = convert_from_bytes(pdf_bytes, dpi=dpi)
    results = []
    for i, img in enumerate(imgs):
        buf = io.BytesIO()
        img.save(buf, format=fmt)
        results.append((f"page_{i+1}.{fmt.lower()}", buf.getvalue()))
    return results

# ── 9. Add text annotation ────────────────────────────────────────

def add_text_annotation(pdf_bytes, text, page_num=1, x=100, y=100,
                         font_size=14, color_hex="#cc0000"):
    annot_buf = io.BytesIO()
    c = rl_canvas.Canvas(annot_buf, pagesize=letter)
    r2 = int(color_hex[1:3],16)/255
    g2 = int(color_hex[3:5],16)/255
    b2 = int(color_hex[5:7],16)/255
    c.setFillColorRGB(r2, g2, b2)
    c.setFont("Helvetica", font_size)
    c.drawString(x, y, text)
    c.save()
    annot_page = PdfReader(io.BytesIO(annot_buf.getvalue())).pages[0]

    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i+1 == page_num:
            page.merge_page(annot_page)
        writer.add_page(page)
    out = io.BytesIO(); writer.write(out); return out.getvalue()

# ── 10. Redact ────────────────────────────────────────────────────

def redact_pdf(pdf_bytes, page_num=1, x=100, y=100, width=200, height=20):
    """Cover region with black rectangle."""
    redact_buf = io.BytesIO()
    c = rl_canvas.Canvas(redact_buf, pagesize=letter)
    c.setFillColorRGB(0, 0, 0)
    c.rect(x, y, width, height, fill=1, stroke=0)
    c.save()
    redact_page = PdfReader(io.BytesIO(redact_buf.getvalue())).pages[0]

    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i+1 == page_num:
            page.merge_page(redact_page)
        writer.add_page(page)
    out = io.BytesIO(); writer.write(out); return out.getvalue()

# ── 11. Reorder pages ─────────────────────────────────────────────

def reorder_pages(pdf_bytes, new_order):
    """new_order: list of 1-indexed page numbers in desired order."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for n in new_order:
        if 1 <= n <= len(reader.pages):
            writer.add_page(reader.pages[n-1])
    out = io.BytesIO(); writer.write(out); return out.getvalue()

# ═══════════════════════════════════════════════════════════════════
#  UI
# ═══════════════════════════════════════════════════════════════════

st.markdown("""
<div class="hero">
    <div class="hero-eyebrow">Professional Document Tools</div>
    <h1>PDF <em>Studio</em></h1>
    <div class="hero-sub">Convert &nbsp;&middot;&nbsp; Merge &nbsp;&middot;&nbsp; Split &nbsp;&middot;&nbsp; Watermark &nbsp;&middot;&nbsp; Compress &nbsp;&middot;&nbsp; Protect &nbsp;&middot;&nbsp; Annotate</div>
</div>
""", unsafe_allow_html=True)

TOOLS = [
    ("📝", "PDF → Word",       "Convert to editable .docx"),
    ("🔗", "Merge PDFs",       "Combine multiple PDFs"),
    ("✂️",  "Split PDF",        "Extract page ranges"),
    ("🔄", "Rotate Pages",     "Rotate any page"),
    ("💧", "Watermark",        "Stamp text on pages"),
    ("🗜️", "Compress",         "Reduce file size"),
    ("🔒", "Protect / Unlock", "Password management"),
    ("🖼️", "Extract Images",   "Save pages as PNG/JPG"),
    ("✍️",  "Add Text",         "Annotate pages"),
    ("⬛", "Redact",           "Black-out content"),
    ("📋", "Reorder Pages",    "Drag pages into order"),
]

# Sidebar tool selector
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">PDF <em>Studio</em></div>
    <div class="sidebar-section">Select Tool</div>
    """, unsafe_allow_html=True)
    tool_names = [t[1] for t in TOOLS]
    selected_tool = st.radio("", tool_names, label_visibility="collapsed")
    st.markdown("---")
    st.markdown("<p style='font-size:0.74rem; color:#3a3028; line-height:1.7;'>Choose a tool, then upload your PDF file to begin.</p>", unsafe_allow_html=True)

icon = [t[0] for t in TOOLS if t[1]==selected_tool][0]
desc = [t[2] for t in TOOLS if t[1]==selected_tool][0]
st.markdown(f'''<div class="tool-heading">
    <span class="tool-icon">{icon}</span>
    <h2>{selected_tool}</h2>
    <span class="tool-tag">{desc}</span>
</div>''', unsafe_allow_html=True)

# ── PDF → Word ────────────────────────────────────────────────────
if selected_tool == "PDF → Word":
    col1, col2 = st.columns([2,1])
    with col1:
        files = st.file_uploader("Upload PDF(s)", type=["pdf"], accept_multiple_files=True)
    with col2:
        font_size = st.slider("Font size", 9, 14, 11)
        include_images = st.checkbox("Extract figures", value=True)
        ocr = st.checkbox("OCR for scanned PDFs", value=True)
        dpi = st.select_slider("Image DPI", [72,100,150,200], value=150)
        as_zip = st.checkbox("Download as ZIP", value=False)

    if files and st.button("Convert to Word"):
        results = []
        bar = st.progress(0)
        for idx, f in enumerate(files):
            with st.spinner(f"Converting {f.name}…"):
                try:
                    docx_bytes = convert_pdf_to_docx(f.read(), font_size, include_images, ocr, dpi)
                    out_name = os.path.splitext(f.name)[0] + ".docx"
                    results.append({"name":f.name,"out":out_name,"bytes":docx_bytes,"ok":True})
                except Exception as e:
                    results.append({"name":f.name,"error":str(e),"ok":False})
            bar.progress((idx+1)/len(files))
        bar.empty()

        if as_zip and any(r["ok"] for r in results):
            zb = io.BytesIO()
            with zipfile.ZipFile(zb,'w',zipfile.ZIP_DEFLATED) as zf:
                for r in results:
                    if r["ok"]: zf.writestr(r["out"], r["bytes"])
            st.download_button("⬇ Download ZIP", zb.getvalue(), "converted.zip", "application/zip")
        else:
            for r in results:
                status = "✅ Converted" if r["ok"] else f"❌ {r.get('error','')}"
                st.markdown(f'<div class="result-card"><div class="fname">📄 {r["name"]}</div><div class="fmeta">{status}</div></div>', unsafe_allow_html=True)
                if r["ok"]:
                    st.download_button(f"⬇ Download {r['out']}", r["bytes"], r["out"],
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=r["out"])

# ── Merge ─────────────────────────────────────────────────────────
elif selected_tool == "Merge PDFs":
    files = st.file_uploader("Upload PDFs to merge (in order)", type=["pdf"], accept_multiple_files=True)
    out_name = st.text_input("Output filename", value="merged.pdf")
    if files:
        st.info(f"{len(files)} files selected — they will be merged in the order shown above.")
        if st.button("Merge PDFs"):
            with st.spinner("Merging…"):
                try:
                    result = merge_pdfs([f.read() for f in files])
                    info = get_pdf_info(result)
                    st.success(f"✅ Merged {len(files)} files → {info['pages']} pages")
                    st.download_button("⬇ Download Merged PDF", result, out_name, "application/pdf")
                except Exception as e:
                    st.error(f"❌ {e}")

# ── Split ─────────────────────────────────────────────────────────
elif selected_tool == "Split PDF":
    f = st.file_uploader("Upload PDF", type=["pdf"])
    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)
        st.info(f"📄 {info['pages']} pages total")

        split_mode = st.radio("Split mode", ["Every page (individual files)", "Custom ranges"])
        if split_mode == "Every page (individual files)":
            ranges = [(i,i) for i in range(1, info['pages']+1)]
        else:
            range_input = st.text_input("Page ranges (e.g. 1-3, 4-6, 7-10)", value=f"1-{info['pages']}")
            ranges = []
            for part in range_input.split(','):
                part = part.strip()
                if '-' in part:
                    a,b = part.split('-')
                    ranges.append((int(a.strip()), int(b.strip())))
                elif part.isdigit():
                    ranges.append((int(part), int(part)))

        if st.button("Split PDF"):
            with st.spinner("Splitting…"):
                try:
                    results = split_pdf(pdf_bytes, ranges)
                    st.success(f"✅ Created {len(results)} file(s)")
                    if len(results) == 1:
                        st.download_button(f"⬇ Download {results[0][0]}", results[0][1], results[0][0], "application/pdf")
                    else:
                        zb = io.BytesIO()
                        with zipfile.ZipFile(zb,'w') as zf:
                            for name, data in results: zf.writestr(name, data)
                        st.download_button("⬇ Download All as ZIP", zb.getvalue(), "split_pages.zip", "application/zip")
                except Exception as e:
                    st.error(f"❌ {e}")

# ── Rotate ────────────────────────────────────────────────────────
elif selected_tool == "Rotate Pages":
    f = st.file_uploader("Upload PDF", type=["pdf"])
    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)
        st.info(f"📄 {info['pages']} pages")
        col1, col2 = st.columns(2)
        with col1:
            angle = st.selectbox("Rotation angle", [90, 180, 270], format_func=lambda x: f"{x}°")
        with col2:
            page_mode = st.radio("Apply to", ["All pages", "Specific pages"])
        page_nums = None
        if page_mode == "Specific pages":
            pg_input = st.text_input("Page numbers (e.g. 1, 3, 5)", value="1")
            page_nums = [int(p.strip()) for p in pg_input.split(',') if p.strip().isdigit()]

        if st.button("Rotate"):
            with st.spinner("Rotating…"):
                try:
                    result = rotate_pdf(pdf_bytes, angle, page_nums)
                    st.success("✅ Done!")
                    out_name = os.path.splitext(f.name)[0] + f"_rotated.pdf"
                    st.download_button("⬇ Download", result, out_name, "application/pdf")
                except Exception as e:
                    st.error(f"❌ {e}")

# ── Watermark ─────────────────────────────────────────────────────
elif selected_tool == "Watermark":
    f = st.file_uploader("Upload PDF", type=["pdf"])
    if f:
        col1, col2 = st.columns(2)
        with col1:
            wm_text = st.text_input("Watermark text", value="CONFIDENTIAL")
            font_size_wm = st.slider("Font size", 20, 100, 60)
            angle = st.slider("Angle (degrees)", 0, 90, 45)
        with col2:
            opacity = st.slider("Opacity", 0.05, 0.8, 0.25)
            color = st.color_picker("Colour", "#cc3333")

        if st.button("Add Watermark"):
            with st.spinner("Adding watermark…"):
                try:
                    result = add_watermark(f.read(), wm_text, opacity, font_size_wm, color, angle)
                    st.success("✅ Watermark added!")
                    out_name = os.path.splitext(f.name)[0] + "_watermarked.pdf"
                    st.download_button("⬇ Download", result, out_name, "application/pdf")
                except Exception as e:
                    st.error(f"❌ {e}")

# ── Compress ──────────────────────────────────────────────────────
elif selected_tool == "Compress":
    files = st.file_uploader("Upload PDF(s)", type=["pdf"], accept_multiple_files=True)
    if files and st.button("Compress"):
        bar = st.progress(0)
        for idx, f in enumerate(files):
            with st.spinner(f"Compressing {f.name}…"):
                try:
                    original = f.read()
                    compressed = compress_pdf(original)
                    saving = (1 - len(compressed)/len(original)) * 100
                    st.markdown(f'<div class="result-card"><div class="fname">🗜️ {f.name}</div><div class="fmeta">{len(original):,} bytes → {len(compressed):,} bytes ({saving:.1f}% smaller)</div></div>', unsafe_allow_html=True)
                    out_name = os.path.splitext(f.name)[0] + "_compressed.pdf"
                    st.download_button(f"⬇ Download {out_name}", compressed, out_name, "application/pdf", key=out_name)
                except Exception as e:
                    st.error(f"❌ {f.name}: {e}")
            bar.progress((idx+1)/len(files))
        bar.empty()

# ── Protect / Unlock ──────────────────────────────────────────────
elif selected_tool == "Protect / Unlock":
    f = st.file_uploader("Upload PDF", type=["pdf"])
    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)
        action = st.radio("Action", ["🔒 Add Password", "🔓 Remove Password"] if info["encrypted"] else ["🔒 Add Password"])

        if "Add Password" in action:
            col1, col2 = st.columns(2)
            with col1: user_pw = st.text_input("User password (to open)", type="password")
            with col2: owner_pw = st.text_input("Owner password (optional)", type="password")
            if st.button("Protect PDF") and user_pw:
                with st.spinner("Protecting…"):
                    try:
                        result = protect_pdf(pdf_bytes, user_pw, owner_pw or None)
                        st.success("✅ PDF is now password protected!")
                        out_name = os.path.splitext(f.name)[0] + "_protected.pdf"
                        st.download_button("⬇ Download", result, out_name, "application/pdf")
                    except Exception as e:
                        st.error(f"❌ {e}")
        else:
            pw = st.text_input("Enter current password", type="password")
            if st.button("Remove Password") and pw:
                with st.spinner("Unlocking…"):
                    try:
                        result = unlock_pdf(pdf_bytes, pw)
                        st.success("✅ Password removed!")
                        out_name = os.path.splitext(f.name)[0] + "_unlocked.pdf"
                        st.download_button("⬇ Download", result, out_name, "application/pdf")
                    except Exception as e:
                        st.error(f"❌ Wrong password or error: {e}")

# ── Extract Images ────────────────────────────────────────────────
elif selected_tool == "Extract Images":
    f = st.file_uploader("Upload PDF", type=["pdf"])
    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)
        col1, col2, col3 = st.columns(3)
        with col1: dpi = st.select_slider("DPI quality", [72,100,150,200,300], value=150)
        with col2: fmt = st.selectbox("Format", ["PNG","JPEG"])
        with col3: st.write(f"**{info['pages']} pages** to extract")

        if st.button("Extract Pages as Images"):
            with st.spinner(f"Rendering {info['pages']} pages…"):
                try:
                    results = extract_as_images(pdf_bytes, dpi, fmt)
                    st.success(f"✅ Extracted {len(results)} images")
                    zb = io.BytesIO()
                    with zipfile.ZipFile(zb,'w') as zf:
                        for name, data in results: zf.writestr(name, data)
                    st.download_button("⬇ Download All as ZIP", zb.getvalue(), "pages.zip", "application/zip")

                    # Preview first 3
                    st.markdown("**Preview (first 3 pages):**")
                    cols = st.columns(min(3, len(results)))
                    for i, (name, data) in enumerate(results[:3]):
                        with cols[i]:
                            st.image(data, caption=name, use_column_width=True)
                except Exception as e:
                    st.error(f"❌ {e}")

# ── Add Text Annotation ───────────────────────────────────────────
elif selected_tool == "Add Text":
    f = st.file_uploader("Upload PDF", type=["pdf"])

    FONTS = {
        "Serif — FreeSerif Regular":      "/usr/share/fonts/truetype/freefont/FreeSerif.ttf",
        "Serif — FreeSerif Bold":         "/usr/share/fonts/truetype/freefont/FreeSerifBold.ttf",
        "Serif — FreeSerif Italic":       "/usr/share/fonts/truetype/freefont/FreeSerifItalic.ttf",
        "Serif — FreeSerif Bold Italic":  "/usr/share/fonts/truetype/freefont/FreeSerifBoldItalic.ttf",
        "Sans — FreeSans Regular":        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        "Sans — FreeSans Bold":           "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        "Sans — FreeSans Oblique":        "/usr/share/fonts/truetype/freefont/FreeSansOblique.ttf",
        "Sans — FreeSans Bold Oblique":   "/usr/share/fonts/truetype/freefont/FreeSansBoldOblique.ttf",
        "Mono — FreeMono":                "/usr/share/fonts/truetype/freefont/FreeMono.ttf",
        "Mono — FreeMono Bold":           "/usr/share/fonts/truetype/freefont/FreeMonoBold.ttf",
        "DejaVu Sans":                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "DejaVu Sans Bold":               "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    }

    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)

        # ── Text & Style controls ──────────────────────────────
        st.markdown("#### Text Content")
        annot_text = st.text_area("Text to add (supports Arabic, English, or mixed)", value="Your text here", height=80)

        col1, col2, col3 = st.columns(3)
        with col1:
            font_name = st.selectbox("Font", list(FONTS.keys()), index=0)
            font_path = FONTS[font_name]
        with col2:
            font_size_a = st.slider("Font size (pt)", 8, 96, 18)
            color_a = st.color_picker("Text colour", "#1c1917")
        with col3:
            page_num = st.number_input("Page", 1, info["pages"], 1)
            opacity_a = st.slider("Opacity", 0.1, 1.0, 1.0)

        st.markdown("---")

        # ── Interactive placement ──────────────────────────────
        st.markdown("#### Position on Page")
        st.caption("Click on the page preview to set where your text will appear, or use the sliders.")

        # Render preview of selected page
        try:
            from pdf2image import convert_from_bytes as cfb
            preview_imgs = cfb(pdf_bytes, dpi=120, first_page=page_num, last_page=page_num)
            preview_img = preview_imgs[0] if preview_imgs else None
        except Exception:
            preview_img = None

        if preview_img:
            prev_w, prev_h = preview_img.size
            # Show preview with text overlay using Pillow
            try:
                from PIL import ImageFont, ImageDraw
                preview_copy = preview_img.copy().convert("RGBA")
                draw = ImageDraw.Draw(preview_copy)
                prev_font = ImageFont.truetype(font_path, max(8, font_size_a))
                # Get coords from sliders (as % of page)
                x_pct = st.slider("Horizontal position (%)", 0, 95, 10, key="xpct")
                y_pct = st.slider("Vertical position (%)", 0, 95, 10, key="ypct")
                px = int(x_pct / 100 * prev_w)
                py = int(y_pct / 100 * prev_h)
                # Parse colour
                ch = color_a.lstrip("#")
                cr, cg, cb = int(ch[0:2],16), int(ch[2:4],16), int(ch[4:6],16)
                alpha = int(opacity_a * 255)
                draw.text((px, py), annot_text, font=prev_font, fill=(cr, cg, cb, alpha))
                # Draw crosshair
                draw.line([(px-10,py),(px+10,py)], fill=(180,80,60,200), width=1)
                draw.line([(px,py-10),(px,py+10)], fill=(180,80,60,200), width=1)
                st.image(preview_copy, caption=f"Preview — Page {page_num}", use_column_width=True)
            except Exception as e:
                st.image(preview_img, caption=f"Page {page_num}", use_column_width=True)
                x_pct = st.slider("Horizontal position (%)", 0, 95, 10, key="xpct2")
                y_pct = st.slider("Vertical position (%)", 0, 95, 10, key="ypct2")
        else:
            x_pct = st.slider("Horizontal position (%)", 0, 95, 10, key="xpct3")
            y_pct = st.slider("Vertical position (%)", 0, 95, 10, key="ypct3")

        st.markdown("---")

        # ── Multiple annotations ───────────────────────────────
        if "annotations" not in st.session_state:
            st.session_state.annotations = []

        cola, colb = st.columns([1,1])
        with cola:
            if st.button("＋ Queue this text annotation"):
                st.session_state.annotations.append({
                    "text": annot_text, "font": font_path,
                    "size": font_size_a, "color": color_a,
                    "opacity": opacity_a, "page": page_num,
                    "x_pct": x_pct, "y_pct": y_pct,
                })
                st.success(f"Queued: \"{annot_text[:30]}…\" on page {page_num}")

        with colb:
            if st.session_state.annotations and st.button("🗑 Clear all queued"):
                st.session_state.annotations = []

        if st.session_state.annotations:
            st.markdown(f"**{len(st.session_state.annotations)} annotation(s) queued:**")
            for i, ann in enumerate(st.session_state.annotations):
                st.markdown(f'<div class="result-card"><div class="fname">"{ann["text"][:40]}"</div>'
                    f'<div class="fmeta">Page {ann["page"]} · {os.path.basename(ann["font"]).replace(".ttf","")} · {ann["size"]}pt · ({ann["x_pct"]}%, {ann["y_pct"]}%)</div></div>',
                    unsafe_allow_html=True)

        if st.button("Apply All Annotations to PDF", use_container_width=True):
            if not st.session_state.annotations:
                st.warning("Queue at least one annotation first.")
            else:
                with st.spinner("Applying annotations…"):
                    try:
                        from PIL import ImageFont, ImageDraw
                        from reportlab.lib.utils import ImageReader
                        from reportlab.pdfgen import canvas as rl_canvas

                        result_bytes = pdf_bytes
                        for ann in st.session_state.annotations:
                            # Render text as transparent image via Pillow (handles Arabic RTL)
                            font_pil = ImageFont.truetype(ann["font"], ann["size"] * 3)
                            dummy = Image.new("RGBA", (1,1))
                            bbox = ImageDraw.Draw(dummy).textbbox((0,0), ann["text"], font=font_pil)
                            tw, th = bbox[2]-bbox[0]+20, bbox[3]-bbox[1]+20
                            txt_img = Image.new("RGBA", (max(1,tw), max(1,th)), (255,255,255,0))
                            ch = ann["color"].lstrip("#")
                            cr,cg,cb = int(ch[0:2],16),int(ch[2:4],16),int(ch[4:6],16)
                            alpha = int(ann["opacity"]*255)
                            ImageDraw.Draw(txt_img).text((10,10), ann["text"], font=font_pil, fill=(cr,cg,cb,alpha))

                            # Get actual PDF page dimensions
                            reader_tmp = PdfReader(io.BytesIO(result_bytes))
                            page_tmp = reader_tmp.pages[ann["page"]-1]
                            pdf_w = float(page_tmp.mediabox.width)
                            pdf_h = float(page_tmp.mediabox.height)

                            # Convert % position to PDF coords (origin bottom-left)
                            x_pdf = (ann["x_pct"] / 100) * pdf_w
                            y_pdf = pdf_h - (ann["y_pct"] / 100) * pdf_h - (th/3)

                            # Build overlay
                            img_buf = io.BytesIO()
                            txt_img.save(img_buf, "PNG"); img_buf.seek(0)
                            overlay_buf = io.BytesIO()
                            c = rl_canvas.Canvas(overlay_buf, pagesize=(pdf_w, pdf_h))
                            c.drawImage(ImageReader(img_buf), x_pdf, y_pdf,
                                        width=tw/3, height=th/3, mask="auto")
                            c.save()

                            # Merge overlay onto selected page
                            overlay_page = PdfReader(io.BytesIO(overlay_buf.getvalue())).pages[0]
                            reader2 = PdfReader(io.BytesIO(result_bytes))
                            writer2 = PdfWriter()
                            for i2, pg in enumerate(reader2.pages):
                                if i2 == ann["page"]-1:
                                    pg.merge_page(overlay_page)
                                writer2.add_page(pg)
                            out2 = io.BytesIO(); writer2.write(out2)
                            result_bytes = out2.getvalue()

                        st.success(f"✅ {len(st.session_state.annotations)} annotation(s) applied!")
                        out_name = os.path.splitext(f.name)[0] + "_annotated.pdf"
                        st.download_button("⬇ Download Annotated PDF", result_bytes, out_name, "application/pdf")
                        st.session_state.annotations = []
                    except Exception as e:
                        st.error(f"❌ {e}")

# ── Redact ────────────────────────────────────────────────────────
elif selected_tool == "Redact":
    f = st.file_uploader("Upload PDF", type=["pdf"])
    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)
        st.info("Position a black redaction box over sensitive content.")
        col1, col2 = st.columns(2)
        with col1:
            page_num_r = st.number_input("Page number", 1, info["pages"], 1)
            x_r = st.slider("X position", 0, 600, 100)
            y_r = st.slider("Y position (from bottom)", 0, 800, 700)
        with col2:
            w_r = st.slider("Width", 10, 500, 200)
            h_r = st.slider("Height", 5, 100, 20)

        if st.button("Apply Redaction"):
            with st.spinner("Redacting…"):
                try:
                    result = redact_pdf(pdf_bytes, page_num_r, x_r, y_r, w_r, h_r)
                    st.success(f"✅ Redacted page {page_num_r}")
                    out_name = os.path.splitext(f.name)[0] + "_redacted.pdf"
                    st.download_button("⬇ Download", result, out_name, "application/pdf")
                except Exception as e:
                    st.error(f"❌ {e}")

# ── Reorder Pages ─────────────────────────────────────────────────
elif selected_tool == "Reorder Pages":
    f = st.file_uploader("Upload PDF", type=["pdf"])
    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)
        st.info(f"📄 {info['pages']} pages — enter the new page order below.")
        default_order = ", ".join(str(i) for i in range(1, info['pages']+1))
        order_input = st.text_input("New page order (comma separated)", value=default_order)
        st.caption("Example: '3, 1, 2' moves page 3 to the front.")

        if st.button("Reorder Pages"):
            with st.spinner("Reordering…"):
                try:
                    new_order = [int(x.strip()) for x in order_input.split(',') if x.strip().isdigit()]
                    result = reorder_pages(pdf_bytes, new_order)
                    st.success(f"✅ Reordered to: {new_order}")
                    out_name = os.path.splitext(f.name)[0] + "_reordered.pdf"
                    st.download_button("⬇ Download", result, out_name, "application/pdf")
                except Exception as e:
                    st.error(f"❌ {e}")
