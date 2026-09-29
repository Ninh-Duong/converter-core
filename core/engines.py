import os
import re
import io
import logging
from pathlib import Path
from core.bootstrap import TESSERACT_CMD, get_gemini_api_key
from core.i18n import t
from core.normalizer import normalize_vietnamese_ocr

# Suppress noisy AFC internal warning from google-genai SDK
logging.getLogger("google_genai.models").setLevel(logging.ERROR)

def is_scanned_pdf(pdf_path: Path) -> bool:
    """Check if a PDF file has an embedded digital text layer or is scanned/image-only."""
    import pymupdf
    doc = pymupdf.open(str(pdf_path))
    total_chars = sum(len(page.get_text().strip()) for page in doc)
    doc.close()
    return total_chars < 50

from dataclasses import dataclass
from html.parser import HTMLParser

@dataclass
class PageContract:
    page_num: int
    char_count: int
    word_count: int
    numbers: set[str]
    structure_keys: set[str]
    has_table: bool

class PageAuditor:
    """Audit rules ensuring page-to-page 1-1 consistency and preventing lost data."""

    @staticmethod
    def extract_pdf_contract(page) -> PageContract:
        text = page.get_text()
        raw_numbers = set(re.findall(r"\b\d+(?:[.,/]\d+)*%?\b", text))
        numbers = {n for n in raw_numbers if len(n) > 1 or n.isdigit()}
        struct_keys = set(re.findall(r"(?:Điều|Khoản|Mục|CHƯƠNG)\s+\d+|[A-D]\.", text, re.IGNORECASE))
        tables = page.find_tables().tables if (hasattr(page, "find_tables") and text.strip()) else []
        has_table = bool(tables) or ("|" in text)
        words = [w for w in text.split() if len(w) > 1]
        return PageContract(
            page_num=page.number + 1,
            char_count=len(text),
            word_count=len(words),
            numbers=numbers,
            structure_keys=struct_keys,
            has_table=has_table
        )

    @staticmethod
    def audit_page(contract: PageContract, output_text: str, docx_tables_count: int = 0) -> tuple[bool, list[str]]:
        errors = []
        if contract.numbers:
            missing_nums = [n for n in contract.numbers if n not in output_text]
            if len(missing_nums) > max(1, int(len(contract.numbers) * 0.15)):
                errors.append(f"Missing numbers: {missing_nums[:5]}")
        if contract.structure_keys:
            missing_keys = [k for k in contract.structure_keys if k.lower() not in output_text.lower()]
            if missing_keys:
                errors.append(f"Missing structural markers: {missing_keys[:5]}")
        if contract.word_count >= 20:
            out_words = len(output_text.split())
            ratio = out_words / contract.word_count
            if ratio < 0.70:
                errors.append(f"Truncated text: {out_words}/{contract.word_count} words ({ratio:.0%})")
        if contract.has_table and docx_tables_count == 0:
            errors.append("PDF page has a table, but output table is missing")

        return len(errors) == 0, errors

class SimpleHTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tables = []
        self._cur_table = []
        self._cur_row = []
        self._cur_cell = []
        self._in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._cur_table = []
        elif tag == "tr":
            self._cur_row = []
        elif tag in ("td", "th"):
            self._cur_cell = []
            self._in_cell = True

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self._cur_row.append("".join(self._cur_cell).strip())
            self._in_cell = False
        elif tag == "tr":
            if self._cur_row:
                self._cur_table.append(self._cur_row)
        elif tag == "table":
            if self._cur_table:
                self.tables.append(self._cur_table)

    def handle_data(self, data):
        if self._in_cell:
            self._cur_cell.append(data)

def add_table_data_to_docx(doc, table_data):
    if not table_data:
        return None
    rows = len(table_data)
    cols = max((len(r) for r in table_data), default=1)
    if rows == 0 or cols == 0:
        return None
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.style = "Table Grid"
    for r_idx, row in enumerate(table_data):
        for c_idx, text in enumerate(row):
            if c_idx < cols:
                cell = tbl.cell(r_idx, c_idx)
                cell.text = text
                if r_idx == 0:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.bold = True
    return tbl

def parse_markdown_table_rows(table_lines):
    parsed = []
    for r in table_lines:
        if re.match(r"^\s*\|?\s*:?-+:?\s*(\|?\s*:?-+:?\s*)+\|?\s*$", r):
            continue
        cols = [c.strip() for c in r.strip("|").split("|")]
        parsed.append(cols)
    return parsed

import time

ACTIVE_MODEL = None

def get_configured_gemini_models() -> list[str]:
    """Return model candidates ordered from newest to fallback, prioritizing user's choice or current active model."""
    global ACTIVE_MODEL
    custom = os.environ.get("GEMINI_MODEL", "").strip()
    
    # Danh sách model theo thứ tự: nhanh/ổn định/quota cao -> dự phòng
    defaults = [
        "gemini-flash-lite-latest",
        "gemini-3.5-flash-lite",
        "gemini-3.1-flash-lite",
        "gemini-flash-latest",
        "gemini-3.8-flash",
        "gemini-3.5-flash",
    ]
    
    # 1. Nếu người dùng chỉ định rõ model trong .env -> ưu tiên số 1
    if custom:
        return [custom] + [m for m in defaults if m != custom]
        
    # 2. Nếu đã có model hoạt động ổn định trước đó -> ưu tiên model đó
    if ACTIVE_MODEL and ACTIVE_MODEL in defaults:
        return [ACTIVE_MODEL] + [m for m in defaults if m != ACTIVE_MODEL]
        
    return defaults

# -------------------------------------------------------------
# 1. Vision AI Engine (Gemini Flash)
# -------------------------------------------------------------
def process_page_with_vision_ai(img_bytes: bytes, api_key: str, feedback: str = "") -> str:
    """Transcribe and structure document image into Markdown and HTML tables with automatic failover."""
    global ACTIVE_MODEL
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key, http_options={"timeout": 60000})
    prompt = (
        "Convert this document page into clean standard format:\n"
        "- 100% ACCURACY: transcribe every single word, number, character, and accent.\n"
        "- NEVER summarize, truncate, or omit any text, notes, or headers.\n"
        "- MULTI-COLUMN ORDER: If the page has 2 columns or 2 side-by-side sections/slides, "
        "read the ENTIRE LEFT column from top to bottom first, then the ENTIRE RIGHT column from top to bottom. "
        "DO NOT interleave lines horizontally across columns.\n"
        "- Use #, ##, ### for titles and headings.\n"
        "- Format ALL tables using standard HTML <table><tr><th>/<td> tags "
        "so multiline cells, headers, and rows are completely preserved without clipping.\n"
        "- Lists: Use - or * for bullets.\n"
        "- Return raw content only (no markdown code block fences, no conversational explanations)."
    )
    if feedback:
        prompt += f"\n\nCRITICAL AUDIT FEEDBACK FROM PREVIOUS ATTEMPT:\n{feedback}\nYou must include ALL missing elements above."

    candidates = get_configured_gemini_models()
    last_err = None

    for idx, model_name in enumerate(candidates):
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[
                        types.Part.from_bytes(data=img_bytes, mime_type="image/png"),
                        prompt
                    ]
                )
                if response and response.text:
                    if ACTIVE_MODEL != model_name:
                        ACTIVE_MODEL = model_name
                    return response.text
            except Exception as e:
                last_err = e
                err_str = str(e)
                
                # Check for transient server overload (503 only; 429 quota will not clear in 2s)
                if ("503" in err_str or "unavailable" in err_str.lower()) and attempt == 0:
                    time.sleep(1)
                    continue

                # Categorize error
                if "503" in err_str or "unavailable" in err_str.lower():
                    reason = "503 Server quá tải"
                elif "504" in err_str or "deadline" in err_str.lower() or "timeout" in err_str.lower():
                    reason = "Timeout / Hết thời gian chờ"
                elif "429" in err_str or "quota" in err_str.lower():
                    reason = "429 Quota limit"
                elif "404" in err_str:
                    reason = "404 Model không khả dụng"
                else:
                    reason = err_str.splitlines()[0][:35]

                next_model = candidates[idx + 1] if idx + 1 < len(candidates) else None
                if next_model:
                    print(f"  [!] Model '{model_name}' gặp sự cố ({reason}). Tự động switch sang '{next_model}'...")
                break

    raise last_err

def render_content_to_docx(raw_text: str, doc) -> int:
    """Render structured text containing HTML tables, Markdown headings, and paragraphs to docx."""
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    clean = re.sub(r"^```(?:html|markdown)?\s*", "", raw_text.strip(), flags=re.IGNORECASE)
    clean = re.sub(r"\s*```$", "", clean)

    chunks = re.split(r"(<table[\s\S]*?</table>)", clean, flags=re.IGNORECASE)
    total_tables = 0

    for chunk in chunks:
        chunk_stripped = chunk.strip()
        if not chunk_stripped:
            continue

        if re.match(r"^<table[\s\S]*?</table>$", chunk_stripped, flags=re.IGNORECASE):
            parser = SimpleHTMLTableParser()
            parser.feed(chunk_stripped)
            for tbl_data in parser.tables:
                if add_table_data_to_docx(doc, tbl_data):
                    total_tables += 1
            continue

        lines = chunk_stripped.split("\n")
        in_md_table = False
        md_table_rows = []

        def flush_md_table():
            nonlocal in_md_table, md_table_rows, total_tables
            if md_table_rows:
                t_data = parse_markdown_table_rows(md_table_rows)
                if add_table_data_to_docx(doc, t_data):
                    total_tables += 1
            md_table_rows = []
            in_md_table = False

        for line in lines:
            stripped = line.strip()
            if not stripped:
                if in_md_table:
                    flush_md_table()
                continue

            if "|" in stripped and (stripped.startswith("|") or stripped.endswith("|") or stripped.count("|") >= 2):
                in_md_table = True
                md_table_rows.append(stripped)
                continue
            elif in_md_table:
                flush_md_table()

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

        if in_md_table:
            flush_md_table()

    return total_tables

def render_markdown_to_docx(md_text: str, doc):
    """Render structured Markdown text into Microsoft Word DOCX elements."""
    render_content_to_docx(md_text, doc)

def get_page_image_slices(page, pdf_in, auto_split_2up: bool = False) -> list[bytes]:
    """Extract page image, keeping full page intact unless 2-up book split is explicitly enabled."""
    from PIL import Image
    pix = page.get_pixmap(dpi=200)
    raw = pix.tobytes("png")
    if auto_split_2up:
        img = Image.open(io.BytesIO(raw))
        w, h = img.size
        if w > h * 1.4:
            slices = []
            for box in ((0, 0, w // 2, h), (w // 2, 0, w, h)):
                buf = io.BytesIO()
                img.crop(box).save(buf, format="PNG")
                slices.append(buf.getvalue())
            return slices
    return [raw]

def save_docx_safely(doc_out, dst: Path) -> Path:
    """Save Word document, safely falling back to <stem>_fixed.docx if target is open/locked in Word."""
    try:
        doc_out.save(str(dst))
        return dst
    except PermissionError:
        fallback = dst.with_name(f"{dst.stem}_fixed{dst.suffix}")
        doc_out.save(str(fallback))
        print(f"\n[!] Warning: '{dst.name}' is open/locked. Saved to '{fallback.name}' instead.")
        return fallback

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

    total_pages = len(pdf_in)
    for idx, page in enumerate(pdf_in, 1):
        print(t("page_progress_ai", page=idx, total=total_pages))
        contract = PageAuditor.extract_pdf_contract(page)
        pix = page.get_pixmap(dpi=200)
        img_bytes = pix.tobytes("png")

        feedback = ""
        output_content = ""
        passed = False
        audit_errors = []

        for attempt in range(2):
            output_content = process_page_with_vision_ai(img_bytes, api_key, feedback=feedback)
            has_table_in_out = bool(re.search(r"<table[\s\S]*?</table>", output_content, re.IGNORECASE)) or ("|" in output_content)
            passed, audit_errors = PageAuditor.audit_page(
                contract,
                output_content,
                docx_tables_count=1 if has_table_in_out else 0
            )
            if passed or not contract.numbers:
                break
            feedback = "; ".join(audit_errors)
            print(f"  [!] Audit retry page {idx} (attempt {attempt+1}): {feedback}")

        tbl_count = render_content_to_docx(output_content, doc_out)

        if passed or not audit_errors:
            print(f"  [✓] Page {idx} Audit PASSED: 1-1 page parity maintained.")
        else:
            print(f"  [!] Page {idx} Audit Warning: {audit_errors[0]}")

        if idx < total_pages:
            doc_out.add_page_break()

    pdf_in.close()
    save_docx_safely(doc_out, dst)

# -------------------------------------------------------------
# 2. Local Tesseract OCR Engine (Offline Fallback)
# -------------------------------------------------------------
def preprocess_page_for_tesseract(pil_img):
    """Nâng cao độ tương phản và làm nét viền ký tự cho Tesseract."""
    from PIL import ImageOps, ImageFilter
    gray = pil_img.convert("L")
    contrast = ImageOps.autocontrast(gray, cutoff=2)
    return contrast.filter(ImageFilter.SHARPEN)

def extract_page_layout_tesseract(img_bytes: bytes, tess_cmd: str):
    import pytesseract
    from PIL import Image
    from pytesseract import Output

    pytesseract.pytesseract.tesseract_cmd = tess_cmd
    raw_img = Image.open(io.BytesIO(img_bytes))
    prep_img = preprocess_page_for_tesseract(raw_img)
    w_total, h_total = prep_img.size
    data = pytesseract.image_to_data(prep_img, lang="vie", config="--psm 1", output_type=Output.DICT)

    raw_lines = []
    cur_key, cur_line = None, None
    for i in range(len(data["text"])):
        w = data["text"][i].strip()
        if not w:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        if key != cur_key:
            cur_key = key
            cur_line = {
                "words": [w],
                "left": data["left"][i],
                "right": data["left"][i] + data["width"][i],
                "top": data["top"][i],
                "bottom": data["top"][i] + data["height"][i],
            }
            raw_lines.append(cur_line)
        else:
            cur_line["words"].append(w)
            cur_line["right"] = max(cur_line["right"], data["left"][i] + data["width"][i])
            cur_line["bottom"] = max(cur_line["bottom"], data["top"][i] + data["height"][i])

    # Detect if page has multi-column/2-up layout (e.g. landscape presentation or 2-column document)
    is_two_col = (w_total > h_total * 1.1) and any(l["left"] > w_total * 0.52 for l in raw_lines) and any(l["left"] < w_total * 0.48 for l in raw_lines)

    for l in raw_lines:
        l["col"] = (1 if l["left"] >= w_total * 0.5 else 0) if is_two_col else 0

    # Sort lines vertically within each column to prevent out-of-order column interleaving
    raw_lines.sort(key=lambda l: (l["col"], round(l["top"] / 10) * 10, l["left"]))

    # Merge line fragments on the same baseline strictly within the same column
    merged = []
    for l in raw_lines:
        if merged and merged[-1]["col"] == l["col"] and abs(merged[-1]["top"] - l["top"]) <= 8:
            merged[-1]["words"].extend(l["words"])
            merged[-1]["right"] = max(merged[-1]["right"], l["right"])
            merged[-1]["bottom"] = max(merged[-1]["bottom"], l["bottom"])
        else:
            merged.append(l)

    page_elements = []
    cur_body = []

    for idx, l in enumerate(merged):
        raw_text = " ".join(l["words"]).strip()
        if not raw_text:
            continue

        # Chuẩn hoá lỗi chính tả quang học tiếng Việt
        text = normalize_vietnamese_ocr(raw_text)

        # Ignore solitary page numbers in footer
        if text.isdigit() and len(text) <= 3 and l["top"] > h_total * 0.85:
            continue

        # Phân loại chính xác các mục, tiêu đề và gạch đầu dòng
        is_bullet = bool(re.match(r"^[-•*+]\s*", text))
        is_title = text.isupper() or text.startswith(("CHƯƠNG", "Điều "))
        is_heading = bool(re.match(r"^(\d+(\.\d+)*|Mục)\b", text))

        if is_title or is_heading or is_bullet:
            if cur_body:
                page_elements.append(("body", " ".join(cur_body)))
                cur_body = []
            if is_bullet:
                clean_bullet = re.sub(r"^[-•*+]\s*", "", text).strip()
                page_elements.append(("bullet", clean_bullet))
            elif is_title:
                page_elements.append(("title", text))
            else:
                page_elements.append(("heading", text))
            continue

        cur_body.append(text)

        next_is_special = False
        if idx + 1 < len(merged):
            nxt = normalize_vietnamese_ocr(" ".join(merged[idx + 1]["words"]).strip())
            if nxt.isupper() or nxt.startswith(("CHƯƠNG", "Điều ")) or re.match(r"^([-•*+]|(\d+(\.\d+)*|Mục)\b)", nxt):
                next_is_special = True

        if next_is_special:
            page_elements.append(("body", " ".join(cur_body)))
            cur_body = []

    if cur_body:
        page_elements.append(("body", " ".join(cur_body)))

    return page_elements

def render_tesseract_elements_to_docx(doc_out, elements):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    for elem_type, text in elements:
        text = text.strip()
        if not text or text in ["]", "[", "|", "~"]:
            continue

        if elem_type == "bullet":
            p = doc_out.add_paragraph(text, style="List Bullet")
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.15
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

    total_pages = len(pdf_in)
    for idx, page in enumerate(pdf_in, 1):
        print(t("page_progress_tess", page=idx, total=total_pages))
        pix = page.get_pixmap(dpi=250)
        elements = extract_page_layout_tesseract(pix.tobytes("png"), TESSERACT_CMD)
        render_tesseract_elements_to_docx(doc_out, elements)
        if idx < total_pages:
            doc_out.add_page_break()

    pdf_in.close()
    save_docx_safely(doc_out, dst)

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
def validate_gemini_connection(api_key: str = None) -> tuple[bool, str]:
    """Validate connection to Gemini API by sending a minimal test ping."""
    global ACTIVE_MODEL
    key = get_gemini_api_key() if api_key is None else api_key
    if not key:
        return False, t("ai_err_no_key")
    try:
        from google import genai
        client = genai.Client(api_key=key, http_options={"timeout": 10000})
        err = None
        for m in get_configured_gemini_models():
            try:
                client.models.generate_content(model=m, contents="ping")
                ACTIVE_MODEL = m
                return True, t("ai_connect_success", model=m)
            except Exception as e:
                err = str(e)
        return False, err or "Connection failed"
    except Exception as e:
        return False, str(e)

def convert_pdf_with_ai(src: Path, dst: Path):
    """Dedicated Gemini Vision AI converter with upfront connection validation."""
    print(t("validating_ai"))
    ok, msg = validate_gemini_connection()
    if not ok:
        print(t("ai_connect_failed", error=msg))
        print(t("falling_back_tesseract"))
        pdf_tesseract_to_docx(src, dst)
        return

    print(msg)
    api_key = get_gemini_api_key()
    try:
        pdf_vision_ai_to_docx(src, dst, api_key)
    except Exception as e:
        print(f"\n[!] Vision AI error: {e}")
        print(t("falling_back_tesseract"))
        pdf_tesseract_to_docx(src, dst)

def convert_pdf_hybrid(src: Path, dst: Path):
    if not is_scanned_pdf(src):
        print(t("detected_vector_pdf"))
        pdf_layout_to_docx(src, dst)
        return

    print(t("detected_scanned_pdf"))
    api_key = get_gemini_api_key()
    if api_key:
        print(t("validating_ai"))
        ok, msg = validate_gemini_connection(api_key)
        if ok:
            print(msg)
            try:
                pdf_vision_ai_to_docx(src, dst, api_key)
                return
            except Exception as e:
                print(f"\n[!] Vision AI error: {e}")
                print(t("falling_back_tesseract"))
        else:
            print(t("ai_connect_failed", error=msg))
            print(t("falling_back_tesseract"))
    else:
        print(t("hint_no_api_key"))

    pdf_tesseract_to_docx(src, dst)

def image_to_docx_ai(src: Path, dst: Path):
    """Convert image to docx via Gemini Vision AI with connection check."""
    print(t("validating_ai"))
    ok, msg = validate_gemini_connection()
    if not ok:
        print(t("ai_connect_failed", error=msg))
        print(t("falling_back_tesseract"))
        image_to_docx_tesseract(src, dst)
        return

    print(msg)
    with open(src, "rb") as f:
        img_bytes = f.read()

    import docx
    from docx.shared import Inches
    print(t("running_vision_ai"))
    doc_out = docx.Document()
    for sec in doc_out.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)
    md_text = process_page_with_vision_ai(img_bytes, get_gemini_api_key())
    render_markdown_to_docx(md_text, doc_out)
    save_docx_safely(doc_out, dst)

def image_to_docx_tesseract(src: Path, dst: Path):
    """Convert image to docx using local offline Tesseract OCR."""
    with open(src, "rb") as f:
        img_bytes = f.read()

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
    save_docx_safely(doc_out, dst)

def docx_to_pdf(src: Path, dst: Path):
    from docx2pdf import convert as d2p
    d2p(str(src), str(dst))

def _print_text_preview(text: str):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    print("\n" + "=" * 60)
    print(f"--- TEXT EXTRACTED FROM IMAGE / NỘI DUNG ĐỌC ĐƯỢC ({len(text.strip())} chars) ---")
    print("-" * 60)
    if lines:
        preview_lines = lines[:12]
        print("\n".join(preview_lines))
        if len(lines) > 12:
            print(f"... ({len(lines) - 12} more lines in file)")
    else:
        print("(No text detected / Không nhận diện được ký tự)")
    print("=" * 60 + "\n")

def image_to_txt_ai(src: Path, dst: Path):
    """Extract all text from image using Gemini Vision AI."""
    print(t("validating_ai"))
    ok, msg = validate_gemini_connection()
    if not ok:
        print(t("ai_connect_failed", error=msg))
        print(t("falling_back_tesseract"))
        image_to_txt_tesseract(src, dst)
        return

    print(msg)
    print(t("running_vision_ai"))
    text = process_page_with_vision_ai(src.read_bytes(), get_gemini_api_key())
    dst.write_text(text, encoding="utf-8")
    _print_text_preview(text)

def preprocess_image_for_tesseract(img):
    """Auto-crop margins, invert dark mode, and add quiet-zone padding for Tesseract OCR."""
    from PIL import ImageOps
    gray = img.convert("L")
    inv = ImageOps.invert(gray)
    bbox = inv.getbbox()
    cropped = gray.crop(bbox) if bbox else gray
    hist = cropped.histogram()
    pixels = sum(hist)
    avg_brightness = sum(i * n for i, n in enumerate(hist)) / pixels if pixels else 255
    if avg_brightness < 128:
        cropped = ImageOps.invert(cropped)
    # Add quiet-zone whitespace border required by Tesseract for edge characters
    return ImageOps.expand(cropped, border=25, fill=255)

def image_to_txt_tesseract(src: Path, dst: Path):
    """Extract all text from image using local Tesseract OCR."""
    if not TESSERACT_CMD:
        print(t("tesseract_not_found"))
        return
    import pytesseract
    from PIL import Image
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    print(t("running_tesseract"))
    raw_img = Image.open(src)
    prep_img = preprocess_image_for_tesseract(raw_img)
    text = pytesseract.image_to_string(prep_img, lang="vie+eng")
    dst.write_text(text, encoding="utf-8")
    _print_text_preview(text)

def get_converter_options(ext: str):
    """Return conversion options available for a file extension."""
    ext = ext.lower()
    if ext == ".pdf":
        return [
            ("mode_pdf_ai", "docx", convert_pdf_with_ai),
            ("mode_pdf_hybrid", "docx", convert_pdf_hybrid),
            ("mode_pdf_tesseract", "docx", pdf_tesseract_to_docx),
        ]
    elif ext in [".docx", ".doc"]:
        return [
            ("mode_docx_pdf", "pdf", docx_to_pdf),
        ]
    elif ext in [".png", ".jpg", ".jpeg"]:
        return [
            ("mode_image_txt_tess", "txt", image_to_txt_tesseract),
            ("mode_image_txt_ai", "txt", image_to_txt_ai),
            ("mode_image_ai", "docx", image_to_docx_ai),
            ("mode_image_tesseract", "docx", image_to_docx_tesseract),
        ]
    return []
