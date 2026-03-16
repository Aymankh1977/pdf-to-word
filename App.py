import streamlit as st
import pdfplumber
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import os

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PDF → Word Converter",
    page_icon="📄",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background: #0f0f0f;
    color: #f0ede8;
}

h1, h2, h3 {
    font-family: 'DM Serif Display', serif !important;
}

/* Header */
.hero {
    text-align: center;
    padding: 3rem 1rem 2rem;
}
.hero h1 {
    font-size: 3rem;
    font-weight: 400;
    color: #f0ede8;
    margin-bottom: 0.4rem;
    letter-spacing: -1px;
}
.hero p {
    color: #888;
    font-size: 1.05rem;
    font-weight: 300;
}
.accent { color: #c8a96e; }

/* Upload area */
[data-testid="stFileUploader"] {
    background: #1a1a1a;
    border: 1.5px dashed #333;
    border-radius: 12px;
    padding: 1rem;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #c8a96e;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    color: #888 !important;
}

/* File pills */
.file-pill {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: #1e1e1e;
    border: 1px solid #2a2a2a;
    border-radius: 8px;
    padding: 8px 14px;
    margin: 4px;
    font-size: 0.88rem;
    color: #ccc;
}

/* Buttons */
.stButton > button {
    background: #c8a96e !important;
    color: #0f0f0f !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.95rem !important;
    padding: 0.6rem 2rem !important;
    transition: opacity 0.2s !important;
    width: 100%;
}
.stButton > button:hover {
    opacity: 0.85 !important;
}

/* Download button */
[data-testid="stDownloadButton"] > button {
    background: #1e1e1e !important;
    color: #c8a96e !important;
    border: 1.5px solid #c8a96e !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
    width: 100%;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #c8a96e !important;
    color: #0f0f0f !important;
}

/* Progress / status */
.stProgress > div > div {
    background: #c8a96e !important;
}
.stSpinner > div {
    border-top-color: #c8a96e !important;
}

/* Result cards */
.result-card {
    background: #1a1a1a;
    border: 1px solid #2a2a2a;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 0.8rem;
}
.result-card .filename {
    font-weight: 500;
    color: #f0ede8;
    margin-bottom: 4px;
}
.result-card .meta {
    font-size: 0.82rem;
    color: #666;
}

/* Success / error badges */
.badge-ok  { color: #6fcf97; font-size: 0.8rem; font-weight: 500; }
.badge-err { color: #eb5757; font-size: 0.8rem; font-weight: 500; }

/* Divider */
hr { border-color: #222 !important; }

/* Hide Streamlit branding */
#MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Conversion helpers ────────────────────────────────────────────────────────

def looks_like_heading(text, font_size=None, is_bold=False):
    stripped = text.strip()
    if not stripped:
        return 0
    if stripped.isupper() and 3 <= len(stripped) <= 80:
        return 1
    if is_bold and len(stripped) <= 80 and not stripped.endswith('.'):
        return 2
    if font_size:
        if font_size >= 18:
            return 1
        if font_size >= 14:
            return 2
        if font_size >= 12 and is_bold:
            return 3
    return 0


def clean_text(text):
    if not text:
        return ""
    text = text.replace('\xa0', ' ')
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def extract_page_content(page):
    content = []
    tables = page.extract_tables()
    try:
        table_bboxes = [t.bbox for t in page.find_tables()]
    except Exception:
        table_bboxes = []

    words = page.extract_words(extra_attrs=["size", "fontname"])
    if not words:
        raw = page.extract_text()
        if raw:
            for line in raw.splitlines():
                line = clean_text(line)
                if line:
                    content.append({"type": "paragraph", "text": line, "font_size": None, "bold": False})
        for tbl in tables:
            if tbl:
                content.append({"type": "table", "data": tbl})
        return content

    lines = {}
    for w in words:
        key = round(w['top'], 1)
        lines.setdefault(key, []).append(w)

    def in_table(y):
        for bbox in table_bboxes:
            if bbox[1] <= y <= bbox[3]:
                return True
        return False

    prev_bottom = None
    for top, line_words in sorted(lines.items()):
        if in_table(top):
            continue
        line_words.sort(key=lambda w: w['x0'])
        line_text = clean_text(' '.join(w['text'] for w in line_words))
        if not line_text:
            continue
        sizes = [w.get('size', 0) for w in line_words if w.get('size')]
        font_size = max(sizes) if sizes else None
        fontnames = [w.get('fontname', '') for w in line_words]
        is_bold = any('Bold' in fn or 'bold' in fn for fn in fontnames)
        gap = (top - prev_bottom) if prev_bottom is not None else 0
        if gap > 12:
            content.append({"type": "blank"})
        content.append({"type": "paragraph", "text": line_text, "font_size": font_size, "bold": is_bold})
        prev_bottom = top + (line_words[0].get('height', 10))

    for tbl in tables:
        if tbl:
            content.append({"type": "table", "data": tbl})

    return content


def add_table_to_doc(doc, table_data):
    rows = [r for r in table_data if any(c for c in r)]
    if not rows:
        return
    num_cols = max(len(r) for r in rows)
    tbl = doc.add_table(rows=len(rows), cols=num_cols)
    tbl.style = 'Table Grid'
    for r_idx, row in enumerate(rows):
        for c_idx, cell_text in enumerate(row):
            if c_idx < num_cols:
                cell = tbl.cell(r_idx, c_idx)
                cell.text = clean_text(cell_text or "")
                if r_idx == 0:
                    for run in cell.paragraphs[0].runs:
                        run.bold = True


def convert_pdf_bytes_to_docx(pdf_bytes: bytes) -> bytes:
    """Convert PDF bytes → DOCX bytes (in-memory, no disk I/O)."""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.2)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            for item in extract_page_content(page):
                if item["type"] == "blank":
                    continue
                elif item["type"] == "paragraph":
                    text = item["text"]
                    level = looks_like_heading(text, item.get("font_size"), item.get("bold", False))
                    if level in (1, 2, 3):
                        doc.add_heading(text, level=level)
                    else:
                        p = doc.add_paragraph(text)
                        if item.get("bold"):
                            for run in p.runs:
                                run.bold = True
                elif item["type"] == "table":
                    add_table_to_doc(doc, item["data"])
            if i < total - 1:
                doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
    <h1>PDF <span class="accent">→</span> Word</h1>
    <p>Upload one or more PDFs and download clean, formatted .docx files instantly.</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drop your PDFs here",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    st.markdown(f"<p style='color:#666; font-size:0.85rem; margin: 0.5rem 0;'>{len(uploaded_files)} file(s) selected</p>", unsafe_allow_html=True)

    if st.button("Convert to Word", use_container_width=True):
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("### Results")

        results = []
        progress = st.progress(0)

        for idx, f in enumerate(uploaded_files):
            with st.spinner(f"Converting {f.name}…"):
                try:
                    docx_bytes = convert_pdf_bytes_to_docx(f.read())
                    out_name = os.path.splitext(f.name)[0] + ".docx"
                    results.append({"name": f.name, "out": out_name, "bytes": docx_bytes, "ok": True})
                except Exception as e:
                    results.append({"name": f.name, "error": str(e), "ok": False})
            progress.progress((idx + 1) / len(uploaded_files))

        progress.empty()

        for r in results:
            with st.container():
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

else:
    st.markdown("""
    <div style='text-align:center; color:#444; padding: 2rem 0; font-size:0.9rem;'>
        Supports text-based PDFs · Tables preserved · Headings detected automatically
    </div>
    """, unsafe_allow_html=True)