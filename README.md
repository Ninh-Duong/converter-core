# Converter Core

A lightweight, self-contained CLI tool for converting documents between **PDF**, **Word (DOCX)**, and **Images** with full layout preservation, table reconstruction, and accurate Vietnamese diacritics support.

---

## Key Features

- **Smart Hybrid Routing**:
  - **Fast Vector Engine (`pdf2docx`)**: Converts digital/native PDFs into Word in ~0.5 seconds, preserving 100% tables, fonts, and alignment without OCR.
  - **Vision AI Engine (Gemini Flash)**: Handles scanned documents, handwritten notes, and photos with 100% accurate Vietnamese diacritics and automatic table formatting.
  - **Offline Tesseract Fallback**: Operates completely offline using native 1:1 image extraction and layout heuristics when no internet or API key is available.
- **Two-Way Document Conversion**: Supports PDF ⇄ DOCX and Image (`.png`, `.jpg`, `.jpeg`) → DOCX.
- **Self-Bootstrapping**: Automatically checks and installs all missing dependencies on the first run with a visual CLI loading progress indicator.
- **Cross-Environment Portability**: All paths are resolved relative to the repository root, ensuring smooth deployment on Windows, Linux, macOS, or Docker.
- **Privacy & Security First**: Pre-configured `.gitignore` ensures your documents, temporary files, and API secrets are never committed to version control.

---

## Project Structure

```text
converter-core/
├── convert.py          # Main CLI application (with auto-installer & hybrid router)
├── requirements.txt    # Python package dependencies
├── .env.example        # Environment variable template
├── .gitignore          # Git exclusion rules (safeguards files/ and .env)
├── tessdata/           # Offline language data (vie, eng, osd)
└── files/              # Workspace directory for input & output files
    └── .gitkeep
```

---

## Prerequisites

- **Python**: Version 3.9 or newer.
- **(Optional) Tesseract OCR**: For offline OCR mode.
  - **Windows**: `winget install UB-Mannheim.TesseractOCR`
  - **Linux**: `sudo apt install tesseract-ocr tesseract-ocr-vie`
  - **macOS**: `brew install tesseract tesseract-lang`

---

## Quick Start

### 1. Clone the repository
```bash
git clone https://github.com/your-username/converter-core.git
cd converter-core
```

### 2. Add files to convert
Place the documents you wish to convert (`.pdf`, `.docx`, `.png`, `.jpg`) into the `files/` folder.

### 3. Run the converter
```bash
python convert.py
```
> **Note**: On the first run, the tool automatically detects missing libraries and installs them. Subsequent runs start instantly.

---

## Configuration (Optional: Vision AI)

For superior accuracy with complex scanned documents and tables, you can enable Google Gemini Flash Vision:

1. Copy the example configuration file:
   ```bash
   cp .env.example .env
   ```
2. Get a free API key from [Google AI Studio](https://aistudio.google.com/app/apikey).
3. Open `.env` and set your key:
   ```env
   GEMINI_API_KEY=your_api_key_here
   ```

If no key is configured, the system automatically falls back to the **local Tesseract OCR engine** with zero interruption.

---

## Usage Guide

1. Run `python convert.py`.
2. The CLI presents an interactive menu listing all documents in `files/`:
   ```text
   === FILE TRONG 'files/' ===
   [1] sample_contract.pdf
   [2] report.docx

   Chọn số file [1-2] (Enter để chọn 1, 'q' để thoát): 1

   Chọn chế độ convert:
   [1] -> .docx (Tự động Hybrid: Fast Vector -> Vision AI -> Tesseract Offline)
   [2] -> .docx (Ép chạy Tesseract OCR Offline)
   Chọn chế độ [1-2] (Enter để chọn 1, 'q' để thoát): 1
   ```
3. The converted document is saved directly in the `files/` directory.

---

## License

This project is licensed under the [MIT License](LICENSE).
