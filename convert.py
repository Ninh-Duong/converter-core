import sys
import os
import subprocess
import importlib.util
import shutil
import re
import io
from pathlib import Path

# Đảm bảo in UTF-8 không lỗi font trên Windows console
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# -------------------------------------------------------------
# 1. TỰ ĐỘNG KIỂM TRA VÀ CÀI ĐẶT THƯ VIỆN KHI CHẠY TRÊN MÔI TRƯỜNG MỚI
# -------------------------------------------------------------
REQUIRED_PACKAGES = [
    ("dotenv", "python-dotenv"),
    ("pymupdf", "pymupdf"),
    ("pdf2docx", "pdf2docx"),
    ("docx2pdf", "docx2pdf"),
    ("docx", "python-docx"),
    ("PIL", "Pillow"),
    ("pytesseract", "pytesseract"),
    ("google.genai", "google-genai"),
]

def ensure_dependencies():
    missing = [pkg for mod, pkg in REQUIRED_PACKAGES if importlib.util.find_spec(mod) is None]
    if missing:
        print("\n" + "=" * 60)
        print("[SETUP] Phát hiện môi trường mới / thiếu thư viện cần thiết.")
        print(f"[SETUP] Đang tự động tải và cài đặt {len(missing)} gói thư viện...")
        print("=" * 60)
        for idx, pkg in enumerate(missing, 1):
            print(f"[{idx}/{len(missing)}] Đang cài đặt '{pkg}'...", end="", flush=True)
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", pkg, "--quiet"],
                    check=True
                )
                print(" [Thành công]")
            except Exception as e:
                print(f" [Lỗi: {e}]")
        print("=" * 60)
        print("[SETUP] Cài đặt hoàn tất! Đang khởi động chương trình...\n")

ensure_dependencies()

from dotenv import load_dotenv

# -------------------------------------------------------------
# 2. ĐƯỜNG DẪN TƯƠNG ĐỐI (DEPLOY ĐƯỢC MỌI MÔI TRƯỜNG)
# -------------------------------------------------------------
REPO_DIR = Path(__file__).parent.resolve()
FILES_DIR = REPO_DIR / "files"
TESSDATA_DIR = REPO_DIR / "tessdata"
ENV_PATH = REPO_DIR / ".env"

# Nạp file .env từ thư mục gốc của repo
if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)

# Thiết lập TESSDATA_PREFIX trỏ về thư mục tessdata tương đối trong repo
if TESSDATA_DIR.exists():
    os.environ["TESSDATA_PREFIX"] = str(TESSDATA_DIR)

# Tìm đường dẫn tesseract.exe linh hoạt
TESSERACT_CMD = shutil.which("tesseract")
if not TESSERACT_CMD and sys.platform == "win32":
    for candidate in [
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        REPO_DIR / "tesseract" / "tesseract.exe",
    ]:
        if candidate.exists():
            TESSERACT_CMD = str(candidate)
            break

def get_gemini_api_key():
    return os.environ.get("GEMINI_API_KEY", "").strip()

def is_scanned_pdf(pdf_path: Path) -> bool:
    """Kiểm tra PDF có chứa text layer thực sự hay là file scan/ảnh."""
    import pymupdf
    doc = pymupdf.open(str(pdf_path))
    total_chars = sum(len(page.get_text().strip()) for page in doc)
    doc.close()
    return total_chars < 50

# ==========================================
# 3. VISION AI PIPELINE (Gemini Flash)
# ==========================================
def process_page_with_vision_ai(img_bytes, api_key):
    """Sử dụng Gemini Flash Vision nhận diện 100% tiếng Việt, bảng biểu và cấu trúc."""
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    prompt = (
        "Hãy chuyển đổi toàn bộ nội dung trong bức ảnh tài liệu này thành định dạng Markdown chuẩn:\n"
        "- Giữ nguyên 100% dấu tiếng Việt và chính tả chính xác.\n"
        "- Đánh dấu tiêu đề bằng # hoặc ## hoặc ###.\n"
        "- Các đoạn văn bản viết liên tục, không ngắt dòng bừa bãi giữa câu.\n"
        "- Bảng biểu hãy định dạng bằng bảng Markdown (| Cột 1 | Cột 2 |).\n"
        "- Chỉ trả về Markdown thuần, không bọc trong ```markdown."
    )

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=img_bytes, mime_type="image/png"),
            prompt
        ]
    )
    return response.text

def render_markdown_to_docx(md_text, doc):
    """Chuyển đổi văn bản Markdown thành các thành phần Word có style chuẩn."""
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

    print("-> Đang xử lý tài liệu bằng Vision AI (Gemini Flash)...")
    pdf_in = pymupdf.open(str(src))
    doc_out = docx.Document()

    for sec in doc_out.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    for idx, page in enumerate(pdf_in, 1):
        print(f"   Trang {idx}/{len(pdf_in)}: Đang phân tích qua Vision AI...")
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

# ==========================================
# 4. LOCAL TESSERACT OCR PIPELINE (Offline Fallback)
# ==========================================
def extract_page_layout_tesseract(img_bytes, tess_cmd):
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
        print("Lỗi: Không tìm thấy Tesseract OCR trên hệ thống. Hãy cài đặt Tesseract để sử dụng chế độ offline.")
        return

    print("-> Đang khởi chạy Tesseract OCR Offline...")
    pdf_in = pymupdf.open(str(src))
    doc_out = docx.Document()

    for sec in doc_out.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    for idx, page in enumerate(pdf_in, 1):
        print(f"   Trang {idx}/{len(pdf_in)} (Tesseract 1:1 layout analysis)...")
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

# ==========================================
# 5. FAST VECTOR PDF PIPELINE (pdf2docx)
# ==========================================
def pdf_layout_to_docx(src: Path, dst: Path):
    from pdf2docx import Converter
    cv = Converter(str(src))
    cv.convert(str(dst))
    cv.close()

# ==========================================
# 6. HYBRID ROUTER
# ==========================================
def convert_pdf_hybrid(src: Path, dst: Path):
    # Tầng 1: PDF văn bản có sẵn text layer -> Chạy pdf2docx (0.5s)
    if not is_scanned_pdf(src):
        print("-> Phát hiện PDF văn bản: Chuyển đổi siêu tốc giữ nguyên 100% bảng & font...")
        pdf_layout_to_docx(src, dst)
        return

    # Tầng 2 & 3: PDF scan / ảnh
    print("-> Phát hiện PDF scan / ảnh chụp (không có text layer).")
    api_key = get_gemini_api_key()
    if api_key:
        pdf_vision_ai_to_docx(src, dst, api_key)
    else:
        print("-> [Gợi ý] Chưa cấu hình GEMINI_API_KEY trong .env. Đang dùng Tesseract OCR Offline...")
        pdf_tesseract_to_docx(src, dst)

def image_to_docx_hybrid(src: Path, dst: Path):
    with open(src, "rb") as f:
        img_bytes = f.read()

    api_key = get_gemini_api_key()
    if api_key:
        import docx
        from docx.shared import Inches
        print("-> Đang nhận diện ảnh bằng Vision AI (Gemini Flash)...")
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
        print("-> Đang nhận diện ảnh bằng Tesseract OCR Offline...")
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

CONVERTERS = {
    ".pdf": [
        ("docx (Tự động Hybrid: Fast Vector -> Vision AI -> Tesseract Offline)", "docx", convert_pdf_hybrid),
        ("docx (Ép chạy Tesseract OCR Offline)", "docx", pdf_tesseract_to_docx),
    ],
    ".docx": [("pdf", "pdf", docx_to_pdf)],
    ".doc":  [("pdf", "pdf", docx_to_pdf)],
    ".png":  [("docx (Hybrid AI/OCR)", "docx", image_to_docx_hybrid)],
    ".jpg":  [("docx (Hybrid AI/OCR)", "docx", image_to_docx_hybrid)],
    ".jpeg": [("docx (Hybrid AI/OCR)", "docx", image_to_docx_hybrid)],
}

def main():
    FILES_DIR.mkdir(exist_ok=True)
    files = [f for f in sorted(FILES_DIR.iterdir()) if f.is_file() and not f.name.startswith("~$") and f.name != ".gitkeep"]

    if not files:
        print(f"Thư mục '{FILES_DIR.name}/' chưa có file. Chép file vào rồi chạy lại.")
        return

    print(f"\n=== FILE TRONG '{FILES_DIR.name}/' ===")
    for i, f in enumerate(files, 1):
        print(f"[{i}] {f.name}")

    while True:
        raw = input(f"\nChọn số file [1-{len(files)}] (Enter để chọn 1, 'q' để thoát): ").strip()
        if raw.lower() == "q":
            return
        if raw == "":
            idx = 0
            break
        if raw.isdigit() and 1 <= int(raw) <= len(files):
            idx = int(raw) - 1
            break
        print(f"Vui lòng nhập số từ 1 đến {len(files)}.")

    src_file = files[idx]
    options = CONVERTERS.get(src_file.suffix.lower())
    if not options:
        print(f"Chưa hỗ trợ định dạng: {src_file.suffix}")
        return

    print("\nChọn chế độ convert:")
    for i, (label, _, _) in enumerate(options, 1):
        print(f"[{i}] -> .{label}")

    while True:
        raw_opt = input(f"Chọn chế độ [1-{len(options)}] (Enter để chọn 1, 'q' để thoát): ").strip()
        if raw_opt.lower() == "q":
            return
        if raw_opt == "":
            opt_idx = 0
            break
        if raw_opt.isdigit() and 1 <= int(raw_opt) <= len(options):
            opt_idx = int(raw_opt) - 1
            break
        print(f"Vui lòng nhập số từ 1 đến {len(options)}.")

    label, target_ext, convert_fn = options[opt_idx]

    base_name = src_file.stem
    if base_name.endswith(".docx") and target_ext == "docx":
        base_name = base_name[:-5]

    dst_file = src_file.parent / f"{base_name}.{target_ext}"
    print(f"\nĐang convert: {src_file.name} -> {dst_file.name}")
    convert_fn(src_file, dst_file)
    print(f"\nXong! File lưu tại: {dst_file}")

if __name__ == "__main__":
    main()
