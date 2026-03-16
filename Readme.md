# PDF → Word Converter

Converts PDF files into clean `.docx` Word documents, preserving:
- Headings (detected by font size and bold styling)
- Paragraphs
- Tables (auto-detected and converted to Word tables)
- Page breaks

---

## Setup (one-time)

**Requirements:** Python 3.8+

```bash
pip install -r requirements.txt
```

---

## Usage

### Convert a single PDF
```bash
python convert.py report.pdf
# Creates: report.docx
```

### Specify a custom output name
```bash
python convert.py report.pdf -o clean_report.docx
```

### Batch convert multiple PDFs
```bash
python convert.py file1.pdf file2.pdf file3.pdf
# Creates: file1.docx, file2.docx, file3.docx
```

### Batch convert all PDFs in a folder (Linux/Mac)
```bash
python convert.py *.pdf
```

---

## Notes

- Works best with **text-based PDFs** (not scanned images).  
  For scanned PDFs, install Tesseract OCR and uncomment the OCR section.
- Tables are detected automatically and formatted with `Table Grid` style.
- Headings are inferred from font size, ALL CAPS, or bold short lines.
- Each PDF page becomes a page break in the Word document.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` |
| Blank output | PDF may be scanned — needs OCR |
| Garbled text | Try a different PDF (some use custom fonts) |