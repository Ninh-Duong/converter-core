import unittest
import tempfile
from pathlib import Path

import docx
import pymupdf

from core.bootstrap import REPO_DIR, FILES_DIR, TESSDATA_DIR
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
        self.assertGreaterEqual(len(img_opts), 1)
        self.assertEqual(img_opts[0][1], "docx")

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
        """Ensure landscape pages (>1.25 aspect ratio) split into 2 sub-pages and portrait remain 1."""
        from core.engines import get_page_image_slices
        # Landscape document (800 x 400, aspect ratio 2.0)
        doc = pymupdf.open()
        p_land = doc.new_page(width=800, height=400)
        slices_land = get_page_image_slices(p_land, doc)
        self.assertEqual(len(slices_land), 2, "Landscape page must split into 2 slices.")

        # Portrait document (400 x 600, aspect ratio 0.67)
        p_port = doc.new_page(width=400, height=600)
        slices_port = get_page_image_slices(p_port, doc)
        self.assertEqual(len(slices_port), 1, "Portrait page must remain 1 slice.")
        doc.close()

    def test_validate_gemini_connection_empty(self):
        """Test that validation fails gracefully when no API key is provided."""
        from core.engines import validate_gemini_connection
        ok, msg = validate_gemini_connection(api_key="")
        self.assertFalse(ok)
        self.assertIn("GEMINI_API_KEY", msg)

if __name__ == "__main__":
    unittest.main(verbosity=2)

