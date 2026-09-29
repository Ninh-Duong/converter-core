"""
Vietnamese OCR Post-Normalizer
Sửa chữa lỗi chính tả quang học (Optical Confusion Matrix) và chuẩn hoá dấu Unicode cho kết quả OCR tiếng Việt.
"""

import re
import unicodedata

# Bảng tra cứu các lỗi nhầm lẫn quang học kinh điển của Tesseract với tiếng Việt
OCR_REPLACEMENTS = [
    # Cụm từ và từ thông dụng bị gãy dấu thanh / nét
    (r"\bùxc\b", "đức"),
    (r"\b[Đđ]ạo\s+ức\b", "Đạo đức"),
    (r"\bđạu\s+đức\b", "đạo đức"),
    (r"\bnghê\b", "nghề"),
    (r"\bụhe\b", "nghề"),
    (r"\bmghe\b", "nghề"),
    (r"\b[ẸE]h[eêễ]\b", "nghề"),
    (r"\bn[\|\s]+ghe\b", "nghề"),
    (r"\bn\s+[ẸE]he\b", "nghề"),
    (r"\bhã[aáảảnh]+\b", "hành"),
    (r"\bhành\s+ngh\b", "hành nghề"),
    (r"\bnghệ\s+nghiệp\b", "nghề nghiệp"),
    (r"\bnghề\s+nghiện\b", "nghề nghiệp"),
    (r"\b[Uu]̉?mg\s+qu[aá]\s+trình\b", "Trong quá trình"),
    (r"\b[Uu]ong\s+qu[aá]\s+trình\b", "Trong quá trình"),
    (r"\btuần\s+thủ\b", "tuân thủ"),
    (r"\btuần\s+theo\b", "tuân theo"),
    (r"\bphãi\s+tuần\b", "phải tuân"),
    (r"\bbảo\s+đ[aă]m\b", "bảo đảm"),
    (r"\bđột\s+lập\b", "độc lập"),
    (r"\bvăn\s+bắn\b", "văn bản"),
    (r"\bchuân\b", "chuẩn"),
    (r"\bchuẩn\s+mục\b", "chuẩn mực"),
    (r"\bchuân\s+mực\b", "chuẩn mực"),
    (r"\bHiển\s+Map\b", "Hiến pháp"),
    (r"\bHiển\s+pháp\b", "Hiến pháp"),
    (r"\bHiến\s+Map\b", "Hiến pháp"),
    (r"\bHiển\s+mủap\b", "Hiến pháp"),
    (r"\bcông\s+dumg\b", "công chứng"),
    (r"\bcông\s+du['’`]?mg\b", "công chứng"),
    (r"\bcông\s+chúứng\b", "công chứng"),
    (r"\bcũng\s+chứng\b", "công chứng"),
    (r"\bnồng\s+chứng\b", "công chứng"),
    (r"\bdu['’`]?mg\b", "chứng"),
    (r"\bcông\s+chíng\b", "công chứng"),
    (r"\bcông\s+chúng\s+viên\b", "công chứng viên"),
    (r"\b(tổ\s+chức\s+hành\s+nghề|luật|văn\s+phòng|hoạt\s+động)\s+công\s+chúng\b", r"\1 công chứng"),
    (r"\bkhẳng\s+ùnh\b", "khẳng định"),
    (r"\bquy\s+ùnh\b", "quy định"),
    (r"\bquý\s+định\b", "quy định"),
    (r"\bxác\s+ùnh\b", "xác định"),
    (r"\bùnh\b", "định"),
    (r"\bphập\s+luật\b", "pháp luật"),
    (r"\bbỏ\s+nhiệm\b", "bổ nhiệm"),
    (r"\bbô\s+nhiệm\b", "bổ nhiệm"),
    (r"\bhợp\s+đông\b", "hợp đồng"),
    (r"\bhợp\s+[fỡ\W]*[l\W]*[aãâä]*ng\b", "hợp đồng"),
    (r"\bđỏi\b", "đổi"),
    (r"\btô\s+chức\b", "tổ chức"),
    (r"\btốổ\s+chức\b", "tổ chức"),
    (r"\bchát\s+lượng\b", "chất lượng"),
    (r"\bthâm\s+quyền\b", "thẩm quyền"),
    (r"\bthâm\s+[ựjr“ển]+\b", "thẩm quyền"),
    (r"\bcân\s+thiết\b", "cần thiết"),
    (r"\bbát\s+hợp\s+lý\b", "bất hợp lý"),
    (r"\bthuân\s+thục\b", "thuần thục"),
    (r"\bthuận\s+thục\b", "thuần thục"),
    (r"\bgóp\s+phân\b", "góp phần"),
    (r"\bbộ\s+phân\b", "bộ phận"),
    (r"\bđôi\s+hỏi\b", "đòi hỏi"),
    (r"\bthủ\s+lao\b", "thù lao"),
    (r"\bmúc\s+thù\s+lao\b", "mức thù lao"),
    (r"\bquản\s+lý\s+nhà\s+nưóc\b", "quản lý nhà nước"),
    (r"\bnhụ\s+cầu\b", "nhu cầu"),
    (r"\bhoạt\s+ửng\b", "hoạt động"),
    (r"\bhoại\s+động\b", "hoạt động"),
    (r"\bxã\s+ủ[›>|]l\b", "xã hội"),
    (r"\bxã\s+hột\s+hóa\b", "xã hội hóa"),
    (r"\bThời\s+gan\b", "Thời gian"),
    (r"\bThỡi\s+gian\b", "Thời gian"),
    (r"\bxác\s+mình\b", "xác minh"),
    (r"\bxắc\s+mình\b", "xác minh"),
    (r"\bnhắn\s+thân\b", "nhân thân"),
    (r"\bcổ\s+tình\b", "cố tình"),
    (r"\btư\s+pháp\s+bỏ\s+trợ\b", "tư pháp bổ trợ"),
    (r"\bụhể\s+trpháp\b", "nghề tư pháp"),
    (r"\bnghiệp\s+vụ\s+căn\s+kẽ\b", "nghiệp vụ cặn kẽ"),
    (r"\bgiao\s+địch\b", "giao dịch"),
    (r"\bgrao\s+dịch\b", "giao dịch"),
    (r"\bđịch\s+vụ\b", "dịch vụ"),
    (r"\btiền\s+tầng\b", "tiền đề"),
    (r"\bnền\s+tầng\b", "nền tảng"),
    (r"\btống\s+cường\b", "tăng cường"),
    (r"\bđổi\s+với\b", "đối với"),
    (r"\bliệm\s+chính\b", "liêm chính"),
    (r"\btnủ\s+tục\b", "thủ tục"),
    (r"\bphít\s+triển\b", "phát triển"),
    (r"\bphất\s+triển\b", "phát triển"),
    (r"\blạp\s+thời\b", "kịp thời"),
    (r"\bnhần\s+dân\b", "nhân dân"),
    (r"\bnhận\s+dân\b", "nhân dân"),
    (r"\bnước\s+tạ\b", "nước ta"),
    (r"\bđảt\s+nước\b", "đất nước"),
    (r"\bcảng\s+rõ\b", "càng rõ"),
    (r"\bvai\s+trô\b", "vai trò"),
    (r"\bđối\s+sống\b", "đời sống"),
    (r"\bđóp\s+ứng\b", "đáp ứng"),
    (r"\bchuyên\s+môn\s+sấu\b", "chuyên môn sâu"),
    (r"\bbấi\s+động\s+sản\b", "bất động sản"),
    (r"\btính\s+chếi\b", "tính chất"),
    (r"\bvì\s+phạm\b", "vi phạm"),
    (r"\btranh\s+chân\b", "tranh chấp"),
    (r"\btải\s+sân\b", "tài sản"),
    (r"\bkẻ\s+biên\b", "kê biên"),
    (r"\bchuyên\s+rJhu\*?ơng\b", "chuyển nhượng"),
    (r"\bhậu\s+quả\s+phần\s+lý\b", "hậu quả pháp lý"),
    (r"\bdầu\s+hiệu\b", "dấu hiệu"),
    (r"\blãnh\s+mạnh\b", "lành mạnh"),
    (r"\bchuẩn\s+mục\s+đạp\s+đức\b", "chuẩn mực đạo đức"),
    (r"\bthực\s+tiến\b", "thực tiễn"),
    (r"\bchỉ\s+phối\b", "chi phối"),
    (r"\bkháng\s+chỉ\s+cần\b", "không chỉ cần"),
    (r"\bvững\s+vảng\b", "vững vàng"),
    (r"\btích\s+lấy\b", "tích lũy"),
    (r"\bgửi\s+găm\b", "gửi gắm"),
]

def _make_case_replacer(replacement: str):
    def replacer(match: re.Match) -> str:
        orig = match.group(0)
        # If the replacement contains regex group references like \1, let re handle it directly
        if "\\" in replacement:
            return match.expand(replacement)
        if orig.isupper():
            return replacement.upper()
        if orig.islower():
            return replacement.lower()
        if orig[0].isupper() and (len(orig) == 1 or orig[1:].islower()):
            return replacement.capitalize()
        return replacement
    return replacer

def normalize_vietnamese_ocr(text: str) -> str:
    """
    Chuẩn hoá văn bản tiếng Việt sau OCR:
    1. Chuyển đổi về chuẩn Unicode dựng sẵn (NFC).
    2. Sửa các lỗi quang học kinh điển của Tesseract với case-matching.
    3. Làm sạch các ký tự rác đầu/cuối dòng.
    """
    if not text:
        return ""

    # 1. Chuẩn hoá Unicode NFC
    normalized = unicodedata.normalize("NFC", text)

    # 2. Thay thế các lỗi nhầm lẫn quang học với case-matching
    for pattern, replacement in OCR_REPLACEMENTS:
        normalized = re.sub(pattern, _make_case_replacer(replacement), normalized, flags=re.IGNORECASE)

    # 3. Làm sạch ký tự rác OCR phổ biến ở đầu/cuối từ
    # Ví dụ: "'Hoạt động" -> "Hoạt động", "- . Bảo vệ" -> "- Bảo vệ"
    normalized = re.sub(r"^[‘'\"`]\s*", "", normalized)
    normalized = re.sub(r"^[-•*+]\s*\.\s*", "- ", normalized)

    return normalized
