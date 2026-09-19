# Multi-language dictionary (English primary, Vietnamese optional)

TRANSLATIONS = {
    "en": {
        "app_title": "=== DOCUMENT CONVERTER CORE ===",
        "files_in_dir": "=== FILES IN 'files/' ===",
        "no_files": "The 'files/' directory is empty. Place your files there and run again.",
        "select_file": "\nSelect file number [1-{count}] (Enter for 1, 'l' to switch language, 'q' to quit): ",
        "select_mode_title": "\nSelect conversion mode:",
        "select_mode": "Select mode [1-{count}] (Enter for 1, 'q' to quit): ",
        "converting": "\nConverting: {src} -> {dst}",
        "done": "\nDone! Output saved at: {path}",
        "unsupported": "Unsupported format: {ext}",
        "invalid_choice": "Invalid choice. Please enter a number between 1 and {count}.",
        "language_switched": "\nLanguage switched to: {lang}",
        "switch_lang_hint": "Type 'l' to toggle language (English / Tiếng Việt)",
        
        # Engine status messages
        "detected_vector_pdf": "-> Detected native digital PDF (with text layer): Converting directly in ~0.5s...",
        "detected_scanned_pdf": "-> Detected scanned document / image (no text layer).",
        "running_vision_ai": "-> Processing with Vision AI (Gemini Flash) for 100% accuracy...",
        "running_tesseract": "-> Processing with local Tesseract OCR (Offline)...",
        "hint_no_api_key": "-> [Notice] No GEMINI_API_KEY found in .env. Falling back to offline Tesseract OCR...",
        "tesseract_not_found": "Error: Tesseract OCR is not installed. Please install Tesseract for offline OCR.",
        "page_progress_ai": "   Page {page}/{total}: Analyzing with Vision AI...",
        "page_progress_tess": "   Page {page}/{total}: Scanning with Tesseract (1:1 layout)...",
        "validating_ai": "-> Checking Gemini Vision AI connection...",
        "ai_connect_success": "-> [OK] Successfully connected to Gemini Vision AI (gemini-flash-latest)!",
        "ai_connect_failed": "-> [FAIL] Cannot connect to Gemini AI: {error}",
        "falling_back_tesseract": "-> Automatically falling back to local offline Tesseract OCR...",
        "ai_err_no_key": "No GEMINI_API_KEY found in .env or environment.",

        # Mode labels
        "mode_pdf_ai": "docx (Gemini Vision AI - 100% Accuracy, Validates Connection)",
        "mode_pdf_hybrid": "docx (Smart Hybrid: Auto Vector -> Vision AI -> Tesseract Offline)",
        "mode_pdf_tesseract": "docx (Force Local Tesseract OCR Offline, No AI)",
        "mode_docx_pdf": "pdf (Convert Word to PDF)",
        "mode_image_ai": "docx (Gemini Vision AI - Cloud Accuracy)",
        "mode_image_tesseract": "docx (Local Tesseract OCR Offline, No AI)",
        "mode_image_docx": "docx (Smart Hybrid AI / OCR to Word)",
    },
    "vi": {
        "app_title": "=== BỘ CÔNG CỤ CHUYỂN ĐỔI TÀI LIỆU ===",
        "files_in_dir": "=== DANH SÁCH FILE TRONG 'files/' ===",
        "no_files": "Thư mục 'files/' đang trống. Hãy chép file vào đây rồi chạy lại.",
        "select_file": "\nChọn số file [1-{count}] (Enter để chọn 1, 'l' đổi ngôn ngữ, 'q' để thoát): ",
        "select_mode_title": "\nChọn chế độ chuyển đổi:",
        "select_mode": "Chọn chế độ [1-{count}] (Enter để chọn 1, 'q' để thoát): ",
        "converting": "\nĐang chuyển đổi: {src} -> {dst}",
        "done": "\nXong! File lưu tại: {path}",
        "unsupported": "Chưa hỗ trợ định dạng: {ext}",
        "invalid_choice": "Lựa chọn sai. Vui lòng nhập số từ 1 đến {count}.",
        "language_switched": "\nĐã chuyển ngôn ngữ sang: {lang}",
        "switch_lang_hint": "Gõ 'l' để chuyển đổi ngôn ngữ (Tiếng Việt / English)",
        
        # Engine status messages
        "detected_vector_pdf": "-> Phát hiện PDF văn bản: Chuyển đổi siêu tốc trong ~0.5s (giữ nguyên bảng & font)...",
        "detected_scanned_pdf": "-> Phát hiện PDF scan / ảnh chụp (không có lớp text).",
        "running_vision_ai": "-> Đang xử lý bằng Vision AI (Gemini Flash) chuẩn 100% tiếng Việt & bảng biểu...",
        "running_tesseract": "-> Đang xử lý bằng Tesseract OCR Offline (cục bộ)...",
        "hint_no_api_key": "-> [Gợi ý] Chưa có GEMINI_API_KEY trong .env. Tự động dùng Tesseract OCR Offline...",
        "tesseract_not_found": "Lỗi: Không tìm thấy Tesseract OCR. Vui lòng cài đặt Tesseract để dùng chế độ offline.",
        "page_progress_ai": "   Trang {page}/{total}: Đang phân tích qua Vision AI...",
        "page_progress_tess": "   Trang {page}/{total}: Đang quét Tesseract (bố cục 1:1)...",
        "validating_ai": "-> Đang kiểm tra kết nối Gemini Vision AI...",
        "ai_connect_success": "-> [OK] Đã kết nối thành công tới Gemini Vision AI (gemini-flash-latest)!",
        "ai_connect_failed": "-> [THẤT BẠI] Không thể kết nối tới Gemini AI: {error}",
        "falling_back_tesseract": "-> Tự động chuyển về dùng Tesseract OCR Offline cục bộ...",
        "ai_err_no_key": "Không tìm thấy GEMINI_API_KEY trong file .env hoặc biến môi trường.",

        # Mode labels
        "mode_pdf_ai": "docx (Gemini Vision AI - Chuẩn 100% AI Cloud, có kiểm tra kết nối)",
        "mode_pdf_hybrid": "docx (Tự động Hybrid: Fast Vector -> Vision AI -> Tesseract Offline)",
        "mode_pdf_tesseract": "docx (Ép chạy Tesseract OCR Offline, không dùng AI)",
        "mode_docx_pdf": "pdf (Chuyển Word sang PDF)",
        "mode_image_ai": "docx (Gemini Vision AI - Chuẩn AI Cloud)",
        "mode_image_tesseract": "docx (Tesseract OCR Offline cục bộ, không dùng AI)",
        "mode_image_docx": "docx (Hybrid AI / OCR sang Word)",
    }
}

CURRENT_LANGUAGE = "en"

def set_language(lang: str):
    global CURRENT_LANGUAGE
    if lang in TRANSLATIONS:
        CURRENT_LANGUAGE = lang

def get_language() -> str:
    return CURRENT_LANGUAGE

def toggle_language() -> str:
    global CURRENT_LANGUAGE
    CURRENT_LANGUAGE = "vi" if CURRENT_LANGUAGE == "en" else "en"
    return CURRENT_LANGUAGE

def t(key: str, **kwargs) -> str:
    lang = CURRENT_LANGUAGE
    text = TRANSLATIONS.get(lang, {}).get(key) or TRANSLATIONS.get("en", {}).get(key, key)
    if kwargs:
        return text.format(**kwargs)
    return text
