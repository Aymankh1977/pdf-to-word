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
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background-color: #fdf5f0; color: #3a2e2e; }
h1,h2,h3 { font-family: 'DM Serif Display', serif !important; color: #3a2e2e; }

/* Top header */
.hero { text-align: center; padding: 2rem 1rem 1rem; }
.hero h1 { font-size: 2.6rem; font-weight: 400; color: #3a2e2e; letter-spacing: -1px; margin-bottom: 0.2rem; }
.hero p { color: #9a7e7e; font-size: 0.95rem; }
.accent { color: #b5736a; }

/* Tool cards in grid */
.tool-grid { display: flex; flex-wrap: wrap; gap: 12px; margin: 1rem 0; }
.tool-card {
    background: #fdf0ea; border: 1.5px solid #e0c0b8;
    border-radius: 12px; padding: 1rem 1.2rem;
    flex: 1 1 180px; cursor: pointer;
    transition: border-color 0.2s, box-shadow 0.2s;
}
.tool-card:hover { border-color: #b5736a; box-shadow: 0 2px 12px rgba(181,115,106,0.15); }
.tool-card .icon { font-size: 1.6rem; margin-bottom: 6px; }
.tool-card .label { font-weight: 600; font-size: 0.9rem; color: #3a2e2e; }
.tool-card .desc { font-size: 0.78rem; color: #9a7e7e; margin-top: 2px; }

/* File uploader */
[data-testid="stFileUploader"] {
    background: #fdf0ea; border: 1.5px dashed #d4a99e; border-radius: 12px; padding: 0.8rem;
}
[data-testid="stFileUploader"]:hover { border-color: #b5736a; }

/* Buttons */
.stButton > button {
    background: #7a4a45 !important; color: #fdf5f0 !important;
    border: none !important; border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important; font-weight: 500 !important;
    font-size: 0.92rem !important; padding: 0.55rem 1.5rem !important;
}
.stButton > button:hover { opacity: 0.85 !important; }

[data-testid="stDownloadButton"] > button {
    background: #3a2e2e !important; color: #e8c4b8 !important;
    border: 1.5px solid #7a4a45 !important; border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important; font-weight: 500 !important;
    width: 100%;
}
[data-testid="stDownloadButton"] > button:hover { background: #7a4a45 !important; color: #fdf5f0 !important; }

/* Result card */
.result-card {
    background: #fdf0ea; border: 1px solid #e0c0b8;
    border-radius: 12px; padding: 1rem 1.2rem; margin-bottom: 0.6rem;
}
.result-card .fname { font-weight: 600; color: #3a2e2e; }
.result-card .fmeta { font-size: 0.8rem; color: #9a7e7e; margin-top: 2px; }

/* Tabs */
[data-testid="stTabs"] button {
    font-family: 'DM Sans', sans-serif !important;
    color: #9a7e7e !important; font-weight: 500 !important;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #7a4a45 !important; border-bottom-color: #7a4a45 !important;
}

/* Sidebar */
[data-testid="stSidebar"] { background: #f5e6df !important; }
[data-testid="stSidebar"] label { color: #7a4a45 !important; }
[data-testid="stSidebar"] .stMarkdown p { color: #9a7e7e !important; font-size: 0.82rem; }

/* Inputs */
.stTextInput input, .stNumberInput input, .stSelectbox select {
    background: #fdf0ea !important; border-color: #d4a99e !important; color: #3a2e2e !important;
}
.stSlider [data-testid="stThumbValue"] { color: #7a4a45 !important; }
div[data-baseweb="slider"] div { background: #b5736a !important; }

.stProgress > div > div { background: #b5736a !important; }
hr { border-color: #e0c0b8 !important; }
#MainMenu, footer { visibility: hidden; }

/* Page thumbnail */
.page-thumb {
    border: 2px solid #e0c0b8; border-radius: 8px;
    padding: 4px; background: white; text-align: center;
    font-size: 0.75rem; color: #9a7e7e; margin-bottom: 4px;
}
.page-thumb.selected { border-color: #b5736a; }
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
    <h1>PDF <span class="accent">Studio</span></h1>
    <p>Convert · Merge · Split · Rotate · Watermark · Compress · Protect · Annotate</p>
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
    st.markdown("### 🛠 Tools")
    tool_names = [t[1] for t in TOOLS]
    selected_tool = st.radio("", tool_names, label_visibility="collapsed")
    st.markdown("---")
    st.markdown("<p style='color:#9a7e7e;font-size:0.8rem;'>Upload a PDF to get started with any tool.</p>", unsafe_allow_html=True)

st.markdown(f"## {[t[0] for t in TOOLS if t[1]==selected_tool][0]}  {selected_tool}")
st.markdown("---")

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
    if f:
        pdf_bytes = f.read()
        info = get_pdf_info(pdf_bytes)
        col1, col2 = st.columns(2)
        with col1:
            annot_text = st.text_input("Text to add", value="Approved")
            page_num = st.number_input("Page number", 1, info["pages"], 1)
            font_size_a = st.slider("Font size", 8, 72, 16)
        with col2:
            x_pos = st.slider("X position (from left)", 0, 600, 100)
            y_pos = st.slider("Y position (from bottom)", 0, 800, 100)
            color_a = st.color_picker("Text colour", "#cc0000")

        if st.button("Add Text"):
            with st.spinner("Adding text…"):
                try:
                    result = add_text_annotation(pdf_bytes, annot_text, page_num, x_pos, y_pos, font_size_a, color_a)
                    st.success(f"✅ Text added to page {page_num}!")
                    out_name = os.path.splitext(f.name)[0] + "_annotated.pdf"
                    st.download_button("⬇ Download", result, out_name, "application/pdf")
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
