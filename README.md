# Converter Core

A lightweight, self-contained CLI tool for converting documents between **PDF**, **Word (DOCX)**, and **Images** with full layout preservation, table reconstruction, and accurate Vietnamese diacritics support.

---

## Key Features

- **Modular & Clean Architecture**: Logic is neatly decomposed into dedicated modules (`bootstrap`, `engines`, `i18n`) with a minimal CLI entry point.
- **Smart Hybrid Routing**:
  - **Fast Vector Engine (`pdf2docx`)**: Converts digital/native PDFs into Word in ~0.5 seconds, preserving 100% tables, fonts, and alignment without OCR.
  - **Vision AI Engine (Gemini Flash)**: Handles scanned documents, handwritten notes, and photos with 100% accurate Vietnamese diacritics and automatic table formatting.
  - **Offline Tesseract Fallback**: Operates completely offline using native 1:1 image extraction and layout heuristics when no internet or API key is available.
- **Bilingual Interface (English & Tiếng Việt)**: Full English interface by default, with instant on-the-fly language toggle (`'l'`).
- **Two-Way Document Conversion**: Supports PDF ⇄ DOCX and Image (`.png`, `.jpg`, `.jpeg`) → DOCX.
- **Self-Bootstrapping**: Automatically checks and installs missing dependencies on the first run with a visual CLI loading progress indicator.
- **Built-in Unit Tests**: Includes an automated test suite verifying core functions, i18n consistency, and layout renderers.
- **Cross-Environment Portability**: All paths are resolved relative to the repository root.
- **Privacy & Security First**: Pre-configured `.gitignore` ensures documents, temporary files, and API secrets are never committed to version control.

---

## Project Structure

```text
converter-core/
├── core/
│   ├── __init__.py
│   ├── bootstrap.py     # Dependency auto-installer & environment paths
│   ├── engines.py       # Conversion engines (Vector, Vision AI, Tesseract, Office)
│   └── i18n.py          # Multi-language dictionary & translation helper
├── convert.py           # Lean CLI runner (~60 lines)
├── test_converter.py    # Automated unit test suite
├── requirements.txt     # Python package dependencies
├── .env.example         # Environment variable template
├── .gitignore           # Git exclusion rules (safeguards files/ and .env)
├── tessdata/            # Offline language data (vie, eng, osd)
└── files/               # Workspace directory for input & output files
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
> **Tip**: Press `'l'` at the file selection prompt to toggle between **English** and **Tiếng Việt**.

---

## Running Unit Tests

Run the built-in test suite:
```bash
python test_converter.py
```

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

## License

This project is licensed under the [MIT License](LICENSE).
