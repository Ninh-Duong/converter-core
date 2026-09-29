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

def _is_missing(mod):
    try:
        return importlib.util.find_spec(mod) is None
    except ModuleNotFoundError:
        return True

def ensure_dependencies():
    """Auto-check and install missing packages on first run."""
    missing = [pkg for mod, pkg in REQUIRED_PACKAGES if _is_missing(mod)]
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
DOCS_DIR = FILES_DIR / "documents"
IMGS_DIR = FILES_DIR / "images"
TESSDATA_DIR = REPO_DIR / "tessdata"
ENV_PATH = REPO_DIR / ".env"

def init_workspace_dirs():
    """Ensure category directories exist and migrate any files from root files/."""
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    IMGS_DIR.mkdir(parents=True, exist_ok=True)
    if FILES_DIR.exists():
        for item in FILES_DIR.iterdir():
            if item.is_file() and not item.name.startswith("~$") and item.name != ".gitkeep":
                target_dir = IMGS_DIR if item.suffix.lower() in [".png", ".jpg", ".jpeg"] else DOCS_DIR
                target_file = target_dir / item.name
                if not target_file.exists():
                    shutil.move(str(item), str(target_file))

init_workspace_dirs()

if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)
elif (REPO_DIR / ".env.example").exists():
    load_dotenv(dotenv_path=REPO_DIR / ".env.example")

if TESSDATA_DIR.exists():
    os.environ["TESSDATA_PREFIX"] = str(TESSDATA_DIR)

def find_tesseract():
    cmd = shutil.which("tesseract")
    if cmd:
        return cmd
    if sys.platform == "win32":
        for candidate in [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
            REPO_DIR / "tesseract" / "tesseract.exe",
        ]:
            if candidate.exists():
                return str(candidate)
    return None

def ensure_tesseract():
    """Ensure Tesseract OCR binary is available, auto-installing if missing."""
    cmd = find_tesseract()
    if cmd:
        return cmd

    print("\n" + "=" * 60)
    print("[SETUP] First run detected: Missing Tesseract OCR binary.")
    print("[SETUP] Automatically installing Tesseract OCR...")
    print("=" * 60)

    if sys.platform == "win32":
        winget = shutil.which("winget")
        if winget:
            try:
                subprocess.run(
                    [
                        winget, "install", "--id", "UB-Mannheim.TesseractOCR",
                        "--accept-source-agreements", "--accept-package-agreements",
                        "--silent"
                    ],
                    check=True
                )
                cmd = find_tesseract()
                if cmd:
                    print("[SETUP] Tesseract OCR installed successfully!\n")
                    return cmd
            except Exception as e:
                print(f"[SETUP] Winget install notice: {e}")

        # Fallback: Download official UB-Mannheim installer to REPO_DIR / "tesseract"
        try:
            import urllib.request
            installer_url = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe"
            tmp_installer = REPO_DIR / "tesseract_setup.exe"
            target_dir = REPO_DIR / "tesseract"
            print("[SETUP] Downloading standalone Tesseract installer...")
            urllib.request.urlretrieve(installer_url, str(tmp_installer))
            print(f"[SETUP] Unpacking to {target_dir}...")
            subprocess.run([str(tmp_installer), "/S", f"/D={target_dir}"], check=True)
            if tmp_installer.exists():
                tmp_installer.unlink()
            cmd = find_tesseract()
            if cmd:
                print("[SETUP] Tesseract OCR standalone installed successfully!\n")
                return cmd
        except Exception as e:
            print(f"[SETUP] Direct download notice: {e}")

    elif sys.platform.startswith("linux"):
        if shutil.which("apt-get"):
            try:
                subprocess.run(["apt-get", "update", "-qq"], check=True)
                subprocess.run(["apt-get", "install", "-y", "tesseract-ocr", "tesseract-ocr-vie"], check=True)
                return find_tesseract()
            except Exception:
                pass
    elif sys.platform == "darwin":
        if shutil.which("brew"):
            try:
                subprocess.run(["brew", "install", "tesseract", "tesseract-lang"], check=True)
                return find_tesseract()
            except Exception:
                pass

    print("[SETUP] Notice: Could not auto-install Tesseract. Offline OCR may be unavailable.")
    print("=" * 60 + "\n")
    return None

TESSERACT_CMD = ensure_tesseract()

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
