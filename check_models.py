#!/usr/bin/env python3
"""
Check Gemini Models Utility
Liệt kê và kiểm tra các model Gemini khả dụng với API Key của bạn.
"""

import sys
from core.bootstrap import get_gemini_api_key
from core.engines import get_configured_gemini_models

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

def main():
    api_key = get_gemini_api_key()
    if not api_key:
        print("\n[!] Không tìm thấy GEMINI_API_KEY trong file .env.")
        print("[!] Hãy thêm API Key vào file .env trước khi chạy kiểm tra.")
        return

    print("\n" + "=" * 65)
    print("      KIỂM TRA CÁC MODEL GEMINI KHẢ DỤNG VỚI API KEY")
    print("=" * 65)

    try:
        from google import genai
        client = genai.Client(api_key=api_key)
    except Exception as e:
        print(f"[!] Lỗi khởi tạo Google GenAI Client: {e}")
        return

    configured_models = get_configured_gemini_models()
    print(f"[*] API Key: {api_key[:6]}...{api_key[-4:]}")
    print(f"[*] Thứ tự model ưu tiên hiện tại: {', '.join(configured_models[:3])}\n")

    candidates = [
        "gemini-3.8-flash",
        "gemini-3.5-flash",
        "gemini-3-flash-preview",
        "gemini-flash-latest",
        "gemini-flash-lite-latest",
        "gemini-2.5-pro",
        "gemini-pro-latest",
    ]

    working_models = []

    print("-" * 65)
    print(f"{'TÊN MODEL':<30} | {'TRẠNG THÁI':<15} | GHI CHÚ")
    print("-" * 65)

    for m in candidates:
        try:
            res = client.models.generate_content(model=m, contents="ping")
            status = "[✓] Hoạt động"
            note = f"Phản hồi: {res.text.strip()[:15]}"
            working_models.append(m)
        except Exception as e:
            err_msg = str(e)
            if "404" in err_msg:
                status = "[x] Không hỗ trợ"
                note = "Chưa mở hoặc không khả dụng"
            elif "503" in err_msg:
                status = "[!] Tạm quá tải"
                note = "Server Google bận (503)"
            elif "429" in err_msg:
                status = "[!] Hết quota"
                note = "Vượt giới hạn lượt gọi (429)"
            else:
                status = "[x] Lỗi"
                note = err_msg.splitlines()[0][:30]

        print(f"{m:<30} | {status:<15} | {note}")

    print("-" * 65)

    if working_models:
        best_model = working_models[0]
        print(f"\n[KHUYẾN NGHỊ]")
        print(f"Model tối ưu nhất cho tài khoản của bạn: '{best_model}'")
        print(f"Bạn có thể cố định model này bằng cách thêm vào file .env:")
        print(f"\n    GEMINI_MODEL={best_model}\n")
    else:
        print("\n[!] Không có model nào phản hồi thành công. Vui lòng kiểm tra lại API Key hoặc hạn mức tài khoản.")

if __name__ == "__main__":
    main()
