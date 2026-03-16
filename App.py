import io
import re
import os
import zipfile
import streamlit as st
import pdfplumber
from pdf2image import convert_from_bytes
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PDF → Word Converter",
    page_icon="📄",
    layout="centered"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background-color: #2b2b2b; color: #f0ede8; }
h1, h2, h3 { font-family: 'DM Serif Display', serif !important; color: #f0ede8; }

.hero { text-align: center; padding: 2.5rem 1rem 1.5rem; }
.hero h1 { font-size: 2.8rem; font-weight: 400; color: #f0ede8; letter-spacing: -1px; margin-bottom: 0.3rem; }
.hero p { color: #bbb; font-size: 1rem; font-weight: 300; }
.accent { color: #c8a96e; }

[data-testid="stFileUploader"] {
    background: #363636; border: 1.5px dashed #555;
    border-radius: 12px; padding: 1rem;
}
[data-testid="stFileUploader"]:hover { border-color: #c8a96e; }

.stButton > button {
    background: #c8a96e !important; color: #0f0f0f !important;
    border: none !important; border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important; font-size: 0.95rem !important;
    padding: 0.6rem 2rem !important; width: 100%;
}
.stButton > button:hover { opacity: 0.85 !important; }

[data-testid="stDownloadButton"] > button {
    background: #363636 !important; color: #c8a96e !important;
    border: 1.5px solid #c8a96e !important; border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important; width: 100%;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #c8a96e !important; color: #0f0f0f !important;
}

.stProgress > div > div { background: #c8a96e !important; }
.result-card {
    background: #363636; border: 1px solid #484848;
    border-radius: 12px; padding: 1.2rem 1.4rem; margin-bottom: 0.8rem;
}
.result-card .filename { font-weight: 500; color: #f0ede8; margin-bottom: 4px; }
.result-card .meta { font-size: 0.82rem; color: #aaa; }
.option-box {
    background: #363636; border: 1px solid #484848;
    border-radius: 12px; padding: 1.2rem 1.4rem; margin-bottom: 1rem;
}
hr { border-color: #484848 !important; }
#MainMenu, footer { visibility: hidden; }

/* Sidebar */
[data-testid="stSidebar"] { background: #333333 !important; }
[data-testid="stSidebar"] label { color: #ccc !important; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────

def clean_text(text):
    if not text:
        return ""
    text = text.replace('\xa0', ' ')
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def looks_like_heading(text, font_size=None, is_bold=False):
    s = text.strip()
    if not s:
        return 0
    if s.isupper() and 3 <= len(s) <= 80:
        return 1
    if is_bold and len(s) <= 80 and not s.endswith('.'):
        return 2
    if font_size:
        if font_size >= 18: return 1
        if font_size >= 14: return 2
        if font_size >= 12 and is_bold: return 3
    return 0


def set_table_border(table):
    """Add clean borders to a Word table."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'AAAAAA')
        tblBorders.append(border)
    tblPr.append(tblBorders)


def add_table_to_doc(doc, table_data, font_size_pt=11):
    """Add a well-formatted table to the Word doc."""
    rows = [r for r in table_data if any(c for c in r)]
    if not rows:
        return
    num_cols = max(len(r) for r in rows)
    tbl = doc.add_table(rows=len(rows), cols=num_cols)
    tbl.style = 'Table Grid'

    for r_idx, row in enumerate(rows):
        for c_idx in range(num_cols):
            cell_text = row[c_idx] if c_idx < len(row) else ""
            cell = tbl.cell(r_idx, c_idx)
            cell.text = clean_text(cell_text or "")
            para = cell.paragraphs[0]
            for run in para.runs:
                run.font.size = Pt(font_size_pt - 1)
                if r_idx == 0:
                    run.bold = True
                    run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            # Header row shading
            if r_idx == 0:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'E8E8E8')
                tcPr.append(shd)

    set_table_border(tbl)
    doc.add_paragraph()  # spacing after table


def extract_page_content(page):
    """Extract structured content from a PDF page."""
    content = []
    tables = page.extract_tables()
    try:
        table_bboxes = [t.bbox for t in page.find_tables()]
    except Exception:
        table_bboxes = []

    # Extract images / figures with their bounding boxes
    try:
        image_bboxes = [(img['x0'], img['top'], img['x1'], img['bottom'])
                        for img in page.images]
    except Exception:
        image_bboxes = []

    words = page.extract_words(extra_attrs=["size", "fontname"])
    if not words:
        raw = page.extract_text()
        if raw:
            for line in raw.splitlines():
                line = clean_text(line)
                if line:
                    content.append({"type": "paragraph", "text": line,
                                    "font_size": None, "bold": False})
        for tbl in tables:
            if tbl:
                content.append({"type": "table", "data": tbl})
        return content, image_bboxes

    lines = {}
    for w in words:
        lines.setdefault(round(w['top'], 1), []).append(w)

    def in_region(y, bboxes):
        for bbox in bboxes:
            if bbox[1] <= y <= bbox[3]:
                return True
        return False

    prev_bottom = None
    for top, line_words in sorted(lines.items()):
        if in_region(top, table_bboxes):
            continue
        if in_region(top, image_bboxes):
            continue
        line_words.sort(key=lambda w: w['x0'])
        line_text = clean_text(' '.join(w['text'] for w in line_words))
        if not line_text:
            continue
        sizes = [w.get('size', 0) for w in line_words if w.get('size')]
        font_size = max(sizes) if sizes else None
        fontnames = [w.get('fontname', '') for w in line_words]
        is_bold = any('Bold' in fn or 'bold' in fn for fn in fontnames)
        if prev_bottom and (top - prev_bottom) > 12:
            content.append({"type": "blank"})
        content.append({"type": "paragraph", "text": line_text,
                         "font_size": font_size, "bold": is_bold})
        prev_bottom = top + line_words[0].get('height', 10)

    for tbl in tables:
        if tbl:
            content.append({"type": "table", "data": tbl})

    return content, image_bboxes


def crop_image_from_page(page_image: Image.Image, bbox, page_width, page_height):
    """Crop a region from the page image given PDF coordinates."""
    img_w, img_h = page_image.size
    scale_x = img_w / page_width
    scale_y = img_h / page_height
    x0 = int(bbox[0] * scale_x)
    y0 = int(bbox[1] * scale_y)
    x1 = int(bbox[2] * scale_x)
    y1 = int(bbox[3] * scale_y)
    # Add small padding
    pad = 4
    x0 = max(0, x0 - pad)
    y0 = max(0, y0 - pad)
    x1 = min(img_w, x1 + pad)
    y1 = min(img_h, y1 + pad)
    if x1 <= x0 or y1 <= y0:
        return None
    return page_image.crop((x0, y0, x1, y1))


def convert_pdf_to_docx(
    pdf_bytes: bytes,
    font_size: int = 11,
    include_images: bool = True,
    ocr_fallback: bool = True,
    dpi: int = 150,
) -> bytes:
    """Full conversion: text + tables + figures → .docx"""

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.2)

    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(font_size)

    # Pre-render all pages as images for figure extraction
    page_images = []
    if include_images:
        try:
            page_images = convert_from_bytes(pdf_bytes, dpi=dpi)
        except Exception:
            page_images = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages):
            content, image_bboxes = extract_page_content(page)

            # Check if page has extractable text
            raw_text = page.extract_text() or ""
            has_text = len(raw_text.strip()) > 20

            # OCR fallback for scanned pages
            if not has_text and ocr_fallback and page_images:
                try:
                    import pytesseract
                    pg_img = page_images[page_num] if page_num < len(page_images) else None
                    if pg_img:
                        ocr_text = pytesseract.image_to_string(pg_img)
                        for line in ocr_text.splitlines():
                            line = clean_text(line)
                            if line:
                                content.append({"type": "paragraph", "text": line,
                                                "font_size": None, "bold": False})
                except Exception:
                    pass

            # Write content to doc
            for item in content:
                if item["type"] == "blank":
                    continue

                elif item["type"] == "paragraph":
                    text = item["text"]
                    level = looks_like_heading(text, item.get("font_size"), item.get("bold", False))
                    if level in (1, 2, 3):
                        h = doc.add_heading(text, level=level)
                        h.runs[0].font.size = Pt(font_size + (6 - level * 2))
                    else:
                        p = doc.add_paragraph(text)
                        for run in p.runs:
                            run.font.size = Pt(font_size)
                            if item.get("bold"):
                                run.bold = True

                elif item["type"] == "table":
                    add_table_to_doc(doc, item["data"], font_size)

            # Embed figures/images from this page
            if include_images and page_images and page_num < len(page_images):
                pg_img = page_images[page_num]
                # Deduplicate overlapping bboxes
                used = []
                for bbox in image_bboxes:
                    # Skip tiny regions (likely decorative)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                    if w < 20 or h < 20:
                        continue
                    # Skip if overlaps with already-used region
                    overlap = False
                    for u in used:
                        if not (bbox[2] < u[0] or bbox[0] > u[2] or
                                bbox[3] < u[1] or bbox[1] > u[3]):
                            overlap = True
                            break
                    if overlap:
                        continue
                    used.append(bbox)

                    cropped = crop_image_from_page(
                        pg_img, bbox, page.width, page.height
                    )
                    if cropped is None:
                        continue

                    img_buf = io.BytesIO()
                    cropped.save(img_buf, format='PNG')
                    img_buf.seek(0)

                    # Scale to fit page width (max 5 inches)
                    img_w_px, img_h_px = cropped.size
                    max_width = Inches(5)
                    aspect = img_h_px / img_w_px if img_w_px > 0 else 1
                    display_w = min(max_width, Inches(img_w_px / dpi))
                    display_h = display_w * aspect

                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    run.add_picture(img_buf, width=display_w)
                    doc.add_paragraph()  # spacing after image

            # Page break between PDF pages
            if page_num < total - 1:
                doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Sidebar options ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Options")
    font_size = st.slider("Font size (pt)", min_value=9, max_value=14, value=11)
    include_images = st.checkbox("Extract figures & charts", value=True)
    ocr_fallback = st.checkbox("OCR for scanned PDFs", value=True)
    image_dpi = st.select_slider("Image quality (DPI)", options=[72, 100, 150, 200], value=150)
    batch_zip = st.checkbox("Download all as ZIP", value=False)
    st.markdown("---")
    st.markdown("<p style='color:#555; font-size:0.8rem;'>Higher DPI = better quality but slower conversion.</p>", unsafe_allow_html=True)

# ── Main UI ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>PDF <span class="accent">→</span> Word</h1>
    <p>Convert PDFs to clean Word documents · Tables · Figures · Charts · OCR</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drop your PDFs here",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    st.markdown(f"<p style='color:#666; font-size:0.85rem; margin:0.5rem 0;'>{len(uploaded_files)} file(s) selected</p>", unsafe_allow_html=True)

    if st.button("Convert to Word", use_container_width=True):
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("### Results")

        results = []
        progress = st.progress(0)

        for idx, f in enumerate(uploaded_files):
            with st.spinner(f"Converting {f.name}…"):
                try:
                    docx_bytes = convert_pdf_to_docx(
                        f.read(),
                        font_size=font_size,
                        include_images=include_images,
                        ocr_fallback=ocr_fallback,
                        dpi=image_dpi,
                    )
                    out_name = os.path.splitext(f.name)[0] + ".docx"
                    results.append({"name": f.name, "out": out_name,
                                    "bytes": docx_bytes, "ok": True})
                except Exception as e:
                    results.append({"name": f.name, "error": str(e), "ok": False})
            progress.progress((idx + 1) / len(uploaded_files))

        progress.empty()

        # Individual download buttons
        if not batch_zip:
            for r in results:
                st.markdown(f"""
                <div class="result-card">
                    <div class="filename">📄 {r['name']}</div>
                    <div class="meta">{'✓ Converted successfully' if r['ok'] else f'✗ {r.get("error","Unknown error")}'}</div>
                </div>
                """, unsafe_allow_html=True)
                if r["ok"]:
                    st.download_button(
                        label=f"⬇  Download {r['out']}",
                        data=r["bytes"],
                        file_name=r["out"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=r["out"],
                    )

        # ZIP download
        else:
            ok_results = [r for r in results if r["ok"]]
            failed = [r for r in results if not r["ok"]]

            if ok_results:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for r in ok_results:
                        zf.writestr(r["out"], r["bytes"])
                zip_buf.seek(0)

                st.success(f"✅ {len(ok_results)} file(s) converted!")
                st.download_button(
                    label=f"⬇  Download all as ZIP ({len(ok_results)} files)",
                    data=zip_buf.getvalue(),
                    file_name="converted_documents.zip",
                    mime="application/zip",
                    key="zip_download"
                )

            for r in failed:
                st.error(f"❌ {r['name']}: {r.get('error', 'Unknown error')}")

else:
    st.markdown("""
    <div style='text-align:center; color:#888; padding:2rem 0; font-size:0.9rem;'>
        Text · Tables · Figures · Charts · OCR for scanned PDFs
    </div>
    """, unsafe_allow_html=True)
