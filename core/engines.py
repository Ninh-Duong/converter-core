import os
import re
import io
from pathlib import Path
from core.bootstrap import TESSERACT_CMD, get_gemini_api_key
from core.i18n import t

def is_scanned_pdf(pdf_path: Path) -> bool:
    """Check if a PDF file has an embedded digital text layer or is scanned/image-only."""
    import pymupdf
    doc = pymupdf.open(str(pdf_path))
    total_chars = sum(len(page.get_text().strip()) for page in doc)
    doc.close()
    return total_chars < 50

# -------------------------------------------------------------
# 1. Vision AI Engine (Gemini Flash)
# -------------------------------------------------------------
def process_page_with_vision_ai(img_bytes: bytes, api_key: str) -> str:
    """Transcribe and structure document image into Markdown using Gemini Flash Vision."""
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    prompt = (
        "Convert this document page into clean standard Markdown format:\n"
        "- Preserve 100% accurate text, accents, and typography.\n"
        "- Use #, ##, ### for titles and headings.\n"
        "- Do not break sentences across arbitrary newlines; join paragraphs smoothly.\n"
        "- Represent tables as clean Markdown tables (| Col 1 | Col 2 |).\n"
        "- Return raw Markdown only, no code fences."
    )

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=img_bytes, mime_type="image/png"),
            prompt
        ]
    )
    return response.text

def render_markdown_to_docx(md_text: str, doc):
    """Render structured Markdown text into Microsoft Word DOCX elements."""
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    lines = md_text.split("\n")
    in_table = False
    table_rows = []

    def flush_table():
        nonlocal in_table, table_rows
        if not table_rows:
            return
        parsed_rows = []
        for r in table_rows:
            if re.match(r"^\s*\|?\s*:?-+:?\s*(\|?\s*:?-+:?\s*)+\|?\s*$", r):
                continue
            cols = [c.strip() for c in r.strip("|").split("|")]
            parsed_rows.append(cols)
        
        if parsed_rows:
            max_cols = max(len(r) for r in parsed_rows)
            table = doc.add_table(rows=len(parsed_rows), cols=max_cols)
            table.style = "Table Grid"
            for r_idx, row in enumerate(parsed_rows):
                for c_idx, cell_text in enumerate(row):
                    if c_idx < max_cols:
                        cell = table.cell(r_idx, c_idx)
                        cell.text = cell_text
                        if r_idx == 0:
                            for p in cell.paragraphs:
                                for run in p.runs:
                                    run.bold = True
        table_rows = []
        in_table = False

    for line in lines:
        stripped = line.strip()
        if not stripped:
            if in_table:
                flush_table()
            continue

        if stripped.startswith("|") and stripped.endswith("|"):
            in_table = True
            table_rows.append(stripped)
            continue
        elif in_table:
            flush_table()

        if stripped.startswith("# "):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(stripped[2:].strip())
            run.bold = True
            run.font.size = Pt(14)
        elif stripped.startswith("## "):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(stripped[3:].strip())
            run.bold = True
            run.font.size = Pt(12)
        elif stripped.startswith("### "):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(stripped[4:].strip())
            run.bold = True
            run.font.size = Pt(11)
        elif stripped.startswith(("- ", "* ")):
            doc.add_paragraph(stripped[2:].strip(), style="List Bullet")
        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            parts = re.split(r"(\*\*.*?\*\*)", stripped)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)

    if in_table:
        flush_table()

def pdf_vision_ai_to_docx(src: Path, dst: Path, api_key: str):
    import pymupdf, docx
    from docx.shared import Inches

    print(t("running_vision_ai"))
    pdf_in = pymupdf.open(str(src))
    doc_out = docx.Document()

    for sec in doc_out.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    for idx, page in enumerate(pdf_in, 1):
        print(t("page_progress_ai", page=idx, total=len(pdf_in)))
        images = page.get_images()
        if images:
            xref = images[0][0]
            base_img = pdf_in.extract_image(xref)
            img_bytes = base_img["image"]
        else:
            pix = page.get_pixmap(dpi=200)
            img_bytes = pix.tobytes("png")

        md_text = process_page_with_vision_ai(img_bytes, api_key)
        render_markdown_to_docx(md_text, doc_out)
        if idx < len(pdf_in):
            doc_out.add_page_break()

    pdf_in.close()
    doc_out.save(str(dst))

# -------------------------------------------------------------
# 2. Local Tesseract OCR Engine (Offline Fallback)
# -------------------------------------------------------------
def extract_page_layout_tesseract(img_bytes: bytes, tess_cmd: str):
    import pytesseract
    from PIL import Image
    from pytesseract import Output

    pytesseract.pytesseract.tesseract_cmd = tess_cmd
    img = Image.open(io.BytesIO(img_bytes))
    data = pytesseract.image_to_data(img, lang="vie", config="--psm 1", output_type=Output.DICT)

    blocks = {}
    for i in range(len(data["text"])):
        w = data["text"][i].strip()
        if not w:
            continue
        b = data["block_num"][i]
        l = data["line_num"][i]
        if b not in blocks:
            blocks[b] = {}
        if l not in blocks[b]:
            blocks[b][l] = {
                "words": [w],
                "left": data["left"][i],
                "right": data["left"][i] + data["width"][i],
                "top": data["top"][i],
                "bottom": data["top"][i] + data["height"][i],
            }
        else:
            blocks[b][l]["words"].append(w)
            blocks[b][l]["right"] = max(blocks[b][l]["right"], data["left"][i] + data["width"][i])
            blocks[b][l]["bottom"] = max(blocks[b][l]["bottom"], data["top"][i] + data["height"][i])

    page_elements = []

    for b_id in sorted(blocks.keys()):
        lines_dict = blocks[b_id]
        sorted_lines = [lines_dict[l] for l in sorted(lines_dict.keys())]
        if not sorted_lines:
            continue

        max_w = max(l["right"] - l["left"] for l in sorted_lines)
        current_body = []
        for line in sorted_lines:
            text = " ".join(line["words"]).strip()
            if not text:
                continue

            w = line["right"] - line["left"]
            is_short = w < (max_w * 0.85)
            is_title = text.isupper() or text.startswith("CHƯƠNG") or text.startswith("Điều")
            is_heading = text.startswith(("1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "1.1", "1.2", "1.3", "Mục "))

            if is_title or is_heading:
                if current_body:
                    page_elements.append(("body", " ".join(current_body)))
                    current_body = []
                page_elements.append(("title" if is_title else "heading", text))
                continue

            current_body.append(text)
            if is_short and text.endswith((".", ":", ";", "!", "?")):
                page_elements.append(("body", " ".join(current_body)))
                current_body = []

        if current_body:
            page_elements.append(("body", " ".join(current_body)))

    return page_elements

def render_tesseract_elements_to_docx(doc_out, elements):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    for elem_type, text in elements:
        text = text.strip()
        if not text or text in ["]", "[", "|", "~"]:
            continue

        p = doc_out.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.15

        if elem_type == "title":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            run.bold = True
            run.font.size = Pt(13)
        elif elem_type == "heading":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            run.bold = True
            run.font.size = Pt(12)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.add_run(text)
            run.font.size = Pt(12)

def pdf_tesseract_to_docx(src: Path, dst: Path):
    import pymupdf, docx
    from docx.shared import Inches

    if not TESSERACT_CMD:
        print(t("tesseract_not_found"))
        return

    print(t("running_tesseract"))
    pdf_in = pymupdf.open(str(src))
    doc_out = docx.Document()

    for sec in doc_out.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    for idx, page in enumerate(pdf_in, 1):
        print(t("page_progress_tess", page=idx, total=len(pdf_in)))
        images = page.get_images()
        if images:
            xref = images[0][0]
            base_img = pdf_in.extract_image(xref)
            img_bytes = base_img["image"]
        else:
            pix = page.get_pixmap(dpi=200)
            img_bytes = pix.tobytes("png")

        elements = extract_page_layout_tesseract(img_bytes, TESSERACT_CMD)
        render_tesseract_elements_to_docx(doc_out, elements)
        if idx < len(pdf_in):
            doc_out.add_page_break()

    pdf_in.close()
    doc_out.save(str(dst))

# -------------------------------------------------------------
# 3. Fast Vector PDF Engine (pdf2docx)
# -------------------------------------------------------------
def pdf_layout_to_docx(src: Path, dst: Path):
    from pdf2docx import Converter
    cv = Converter(str(src))
    cv.convert(str(dst))
    cv.close()

# -------------------------------------------------------------
# 4. Hybrid Routers & Office Engines
# -------------------------------------------------------------
def convert_pdf_hybrid(src: Path, dst: Path):
    if not is_scanned_pdf(src):
        print(t("detected_vector_pdf"))
        pdf_layout_to_docx(src, dst)
        return

    print(t("detected_scanned_pdf"))
    api_key = get_gemini_api_key()
    if api_key:
        pdf_vision_ai_to_docx(src, dst, api_key)
    else:
        print(t("hint_no_api_key"))
        pdf_tesseract_to_docx(src, dst)

def image_to_docx_hybrid(src: Path, dst: Path):
    with open(src, "rb") as f:
        img_bytes = f.read()

    api_key = get_gemini_api_key()
    if api_key:
        import docx
        from docx.shared import Inches
        print(t("running_vision_ai"))
        doc_out = docx.Document()
        for sec in doc_out.sections:
            sec.top_margin = Inches(1)
            sec.bottom_margin = Inches(1)
            sec.left_margin = Inches(1)
            sec.right_margin = Inches(1)
        md_text = process_page_with_vision_ai(img_bytes, api_key)
        render_markdown_to_docx(md_text, doc_out)
        doc_out.save(str(dst))
    else:
        import docx
        from docx.shared import Inches
        print(t("running_tesseract"))
        doc_out = docx.Document()
        for sec in doc_out.sections:
            sec.top_margin = Inches(1)
            sec.bottom_margin = Inches(1)
            sec.left_margin = Inches(1)
            sec.right_margin = Inches(1)
        elements = extract_page_layout_tesseract(img_bytes, TESSERACT_CMD)
        render_tesseract_elements_to_docx(doc_out, elements)
        doc_out.save(str(dst))

def docx_to_pdf(src: Path, dst: Path):
    from docx2pdf import convert as d2p
    d2p(str(src), str(dst))

def get_converter_options(ext: str):
    """Return conversion options available for a file extension."""
    ext = ext.lower()
    if ext == ".pdf":
        return [
            ("mode_pdf_hybrid", "docx", convert_pdf_hybrid),
            ("mode_pdf_tesseract", "docx", pdf_tesseract_to_docx),
        ]
    elif ext in [".docx", ".doc"]:
        return [
            ("mode_docx_pdf", "pdf", docx_to_pdf),
        ]
    elif ext in [".png", ".jpg", ".jpeg"]:
        return [
            ("mode_image_docx", "docx", image_to_docx_hybrid),
        ]
    return []
