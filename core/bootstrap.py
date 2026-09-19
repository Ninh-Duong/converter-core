import sys
import os
import subprocess
import importlib.util
import shutil
from pathlib import Path

# Đảm bảo UTF-8 trên Windows console
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# Required packages for first-run bootstrapping
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
    """Auto-check and install missing packages on first run."""
    missing = [pkg for mod, pkg in REQUIRED_PACKAGES if importlib.util.find_spec(mod) is None]
    if missing:
        print("\n" + "=" * 60)
        print("[SETUP] First run detected: Missing required dependencies.")
        print(f"[SETUP] Automatically installing {len(missing)} package(s)...")
        print("=" * 60)
        for idx, pkg in enumerate(missing, 1):
            print(f"[{idx}/{len(missing)}] Installing '{pkg}'...", end="", flush=True)
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", pkg, "--quiet"],
                    check=True
                )
                print(" [Done]")
            except Exception as e:
                print(f" [Error: {e}]")
        print("=" * 60)
        print("[SETUP] All dependencies installed successfully!\n")

# Run dependency check immediately
ensure_dependencies()

from dotenv import load_dotenv

# -------------------------------------------------------------
# Path Resolution (Strictly Relative)
# -------------------------------------------------------------
REPO_DIR = Path(__file__).parent.parent.resolve()
FILES_DIR = REPO_DIR / "files"
TESSDATA_DIR = REPO_DIR / "tessdata"
ENV_PATH = REPO_DIR / ".env"

if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)
elif (REPO_DIR / ".env.example").exists():
    load_dotenv(dotenv_path=REPO_DIR / ".env.example")

if TESSDATA_DIR.exists():
    os.environ["TESSDATA_PREFIX"] = str(TESSDATA_DIR)

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
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not key or key == "your_api_key_here":
        for p in [ENV_PATH, REPO_DIR / ".env.example"]:
            if p.exists():
                for line in p.read_text(encoding="utf-8", errors="ignore").splitlines():
                    if line.strip().startswith("GEMINI_API_KEY="):
                        val = line.split("=", 1)[1].strip()
                        if val and val != "your_api_key_here":
                            return val
        return ""
    return key
