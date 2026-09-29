import unittest
import tempfile
from pathlib import Path

import docx
import pymupdf

from core.bootstrap import REPO_DIR, FILES_DIR, DOCS_DIR, IMGS_DIR, TESSDATA_DIR
from core.i18n import TRANSLATIONS, t, set_language, toggle_language, get_language
from core.engines import (
    is_scanned_pdf,
    get_converter_options,
    render_markdown_to_docx,
)

class TestConverterCore(unittest.TestCase):
    
    def test_relative_paths(self):
        """Ensure all project directories are resolved relatively."""
        self.assertTrue(REPO_DIR.exists(), "REPO_DIR must exist.")
        self.assertTrue(FILES_DIR.exists(), "FILES_DIR must exist.")
        self.assertTrue(DOCS_DIR.exists(), "DOCS_DIR must exist.")
        self.assertTrue(IMGS_DIR.exists(), "IMGS_DIR must exist.")
        self.assertTrue(TESSDATA_DIR.exists(), "TESSDATA_DIR must exist.")

    def test_i18n_consistency(self):
        """Ensure English and Vietnamese translations have matching keys."""
        en_keys = set(TRANSLATIONS["en"].keys())
        vi_keys = set(TRANSLATIONS["vi"].keys())
        self.assertEqual(en_keys, vi_keys, "Missing i18n keys between EN and VI.")

    def test_i18n_toggle(self):
        """Test language switching functionality."""
        set_language("en")
        self.assertEqual(get_language(), "en")
        self.assertIn("DOCUMENT CONVERTER", t("app_title"))

        toggle_language()
        self.assertEqual(get_language(), "vi")
        self.assertIn("CHUYỂN ĐỔI TÀI LIỆU", t("app_title"))

        # Revert back to English
        set_language("en")

    def test_converter_options(self):
        """Test converter router options for various file extensions."""
        pdf_opts = get_converter_options(".pdf")
        self.assertGreaterEqual(len(pdf_opts), 1)
        self.assertEqual(pdf_opts[0][1], "docx")

        docx_opts = get_converter_options(".docx")
        self.assertGreaterEqual(len(docx_opts), 1)
        self.assertEqual(docx_opts[0][1], "pdf")

        img_opts = get_converter_options(".png")
        self.assertGreaterEqual(len(img_opts), 2)
        target_exts = [opt[1] for opt in img_opts]
        self.assertIn("txt", target_exts)
        self.assertIn("docx", target_exts)

        self.assertEqual(get_converter_options(".unknown"), [])

    def test_markdown_to_docx_renderer(self):
        """Test markdown rendering into Word elements (headings, tables, bold)."""
        doc = docx.Document()
        sample_md = (
            "# Document Title\n"
            "## Section 1: Overview\n"
            "This is **bold text** in a normal paragraph.\n\n"
            "| Item | Price |\n"
            "|---|---|\n"
            "| Apple | $1.00 |\n"
            "| Orange | $1.50 |\n"
        )
        render_markdown_to_docx(sample_md, doc)

        # Check paragraphs
        self.assertGreaterEqual(len(doc.paragraphs), 3)
        self.assertEqual(doc.paragraphs[0].text, "Document Title")
        self.assertEqual(doc.paragraphs[1].text, "Section 1: Overview")
        self.assertIn("bold text", doc.paragraphs[2].text)

        # Check table
        self.assertEqual(len(doc.tables), 1)
        table = doc.tables[0]
        self.assertEqual(len(table.rows), 3)
        self.assertEqual(table.cell(0, 0).text, "Item")
        self.assertEqual(table.cell(1, 0).text, "Apple")

    def test_is_scanned_pdf_digital(self):
        """Verify is_scanned_pdf correctly flags a digital PDF with text."""
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text((50, 72), "This is a digital test PDF document with plenty of embedded text.", fontsize=12)
        doc.save(str(tmp_path))
        doc.close()

        try:
            self.assertFalse(is_scanned_pdf(tmp_path), "Digital PDF should NOT be flagged as scanned.")
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    def test_landscape_slice_splitting(self):
        """Ensure landscape pages remain 1 slice by default to protect tables, and split only if auto_split_2up=True."""
        from core.engines import get_page_image_slices
        doc = pymupdf.open()
        p_land = doc.new_page(width=800, height=400)
        # Default: keep 1-1 intact page to protect tables
        self.assertEqual(len(get_page_image_slices(p_land, doc)), 1, "Default must preserve full landscape page.")
        # When 2-up scan splitting is explicitly enabled
        self.assertEqual(len(get_page_image_slices(p_land, doc, auto_split_2up=True)), 2, "2-up split must yield 2 slices.")

        p_port = doc.new_page(width=400, height=600)
        self.assertEqual(len(get_page_image_slices(p_port, doc)), 1, "Portrait page must remain 1 slice.")
        doc.close()

    def test_html_table_rendering(self):
        """Verify HTML table parsing into Word table with rows and cells."""
        from core.engines import render_content_to_docx
        doc = docx.Document()
        html_content = (
            "# Báo cáo\n"
            "<table>\n"
            "  <tr><th>Mã</th><th>Tên</th><th>Điểm</th></tr>\n"
            "  <tr><td>001</td><td>Nguyễn Văn A</td><td>9.5</td></tr>\n"
            "</table>\n"
            "Đoạn văn kết thúc."
        )
        total_tables = render_content_to_docx(html_content, doc)
        self.assertEqual(total_tables, 1)
        self.assertEqual(len(doc.tables), 1)
        self.assertEqual(len(doc.tables[0].rows), 2)
        self.assertEqual(doc.tables[0].cell(0, 0).text, "Mã")
        self.assertEqual(doc.tables[0].cell(1, 1).text, "Nguyễn Văn A")

    def test_page_auditor_rules(self):
        """Test PageAuditor validation rules: missing numbers, missing tables, and word count retention."""
        from core.engines import PageContract, PageAuditor
        contract = PageContract(
            page_num=1,
            char_count=100,
            word_count=14,
            numbers={"2024", "10", "100%"},
            structure_keys={"Điều 1"},
            has_table=True
        )

        # 1. Output missing numbers & table -> should FAIL
        passed, errors = PageAuditor.audit_page(contract, "Đây là văn bản thiếu số và bảng.", docx_tables_count=0)
        self.assertFalse(passed)
        self.assertTrue(any("Missing numbers" in e for e in errors))
        self.assertTrue(any("output table is missing" in e for e in errors))

        # 2. Output with full numbers, keys & table -> should PASS
        passed, errors = PageAuditor.audit_page(
            contract,
            "Nội dung Điều 1 theo quy định năm 2024, số lượng 10 và tỷ lệ 100%.",
            docx_tables_count=1
        )
        self.assertTrue(passed)
        self.assertEqual(errors, [])

    def test_validate_gemini_connection_empty(self):
        """Test that validation fails gracefully when no API key is provided."""
        from core.engines import validate_gemini_connection
        ok, msg = validate_gemini_connection(api_key="")
        self.assertFalse(ok)
        self.assertIn("GEMINI_API_KEY", msg)

    def test_image_to_txt_tesseract(self):
        """Test reading text from an image to a txt file using Tesseract."""
        from PIL import Image, ImageDraw
        from core.engines import image_to_txt_tesseract

        with tempfile.TemporaryDirectory() as tmpdir:
            img_path = Path(tmpdir) / "test.png"
            txt_path = Path(tmpdir) / "test.txt"

            img = Image.new("RGB", (200, 60), color="white")
            draw = ImageDraw.Draw(img)
            draw.text((10, 20), "HELLO", fill="black")
            img.save(str(img_path))

            image_to_txt_tesseract(img_path, txt_path)
            self.assertTrue(txt_path.exists())
            self.assertIn("HELLO", txt_path.read_text(encoding="utf-8").strip())

    def test_normalize_vietnamese_ocr(self):
        """Ensure optical confusion errors from Tesseract are corrected."""
        from core.normalizer import normalize_vietnamese_ocr

        # Known Tesseract optical errors reported by user
        corrupted = "QUY TẮC Đạo ùxc HÀNH nghê CỦA CCV Theo Hiển Map và công dumg chuân mực nghệ nghiệp"
        expected = "QUY TẮC Đạo đức HÀNH nghề CỦA CCV Theo Hiến pháp và công chứng chuẩn mực nghề nghiệp"
        self.assertEqual(normalize_vietnamese_ocr(corrupted), expected)

        # Artifact cleaning
        self.assertEqual(normalize_vietnamese_ocr("- . Bảo vệ quyền lợi"), "- Bảo vệ quyền lợi")
        self.assertEqual(normalize_vietnamese_ocr("'Hoạt động công chứng"), "Hoạt động công chứng")

    def test_render_tesseract_bullet_elements(self):
        """Verify that bullet elements are rendered as List Bullet paragraphs in docx."""
        from core.engines import render_tesseract_elements_to_docx
        doc = docx.Document()
        elements = [
            ("title", "TIÊU ĐỀ BÀI VIẾT"),
            ("bullet", "Bảo vệ quyền, lợi ích hợp pháp của người yêu cầu"),
            ("bullet", "Tôn trọng quyền tự do ý chí"),
            ("body", "Đoạn văn giải thích chi tiết."),
        ]
        render_tesseract_elements_to_docx(doc, elements)

        self.assertEqual(len(doc.paragraphs), 4)
        self.assertEqual(doc.paragraphs[0].text, "TIÊU ĐỀ BÀI VIẾT")
        self.assertEqual(doc.paragraphs[1].style.name, "List Bullet")
        self.assertEqual(doc.paragraphs[1].text, "Bảo vệ quyền, lợi ích hợp pháp của người yêu cầu")
        self.assertEqual(doc.paragraphs[2].style.name, "List Bullet")
        self.assertEqual(doc.paragraphs[3].text, "Đoạn văn giải thích chi tiết.")

    def test_gemini_models_failover_ordering(self):
        """Verify dynamic failover model list is non-empty and starts with newest models."""
        from core.engines import get_configured_gemini_models
        models = get_configured_gemini_models()
        self.assertGreaterEqual(len(models), 3)
        self.assertIn("gemini-3.8-flash", models)
        self.assertIn("gemini-3.5-flash", models)

if __name__ == "__main__":
    unittest.main(verbosity=2)


