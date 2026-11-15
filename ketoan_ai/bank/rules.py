"""Bộ nhận diện nội dung chuyển khoản theo quy tắc (không cần AI).

Nội dung ủy nhiệm chi của công ty thường có dạng:
    HCIC TT <mục đích> [CT <công trình> | VP] CHO CTY <nhà cung cấp>-DDMMYY-HH:MM:SS <số ref>
Bộ này tách: mục đích, công trình, nhà cung cấp; tra danh mục; sinh diễn giải tiếng Việt có dấu.
Khi có API key, kết quả này được đưa vào prompt làm gợi ý; khi không có, nó là kết quả chính.
"""
from __future__ import annotations

import re

from .. import utils

# Từ điển phục hồi dấu cho các cụm hay gặp trong nội dung chuyển khoản (chữ hoa không dấu -> có dấu)
TU_DIEN = [
    ("CONG NO", "công nợ"), ("BE TONG", "bê tông"), ("THEP", "thép"), ("VAT TU", "vật tư"), ("TON", "tôn"), ("CAP", "cáp"),
    ("SUA MAY TINH", "sửa máy tính"), ("MAY TINH", "máy tính"), ("MAY PHOTO", "máy photo"), ("MAY KHOAN", "máy khoan"),
    ("NUOC UONG", "nước uống"), ("PHI CONG CHUNG", "phí công chứng"), ("THE CHAP", "thế chấp"), ("CAT", "cát"), ("GACH", "gạch"),
    ("XI MANG", "xi măng"), ("NHAN CONG", "nhân công"), ("VAN CHUYEN", "vận chuyển"), ("LUONG", "lương"), ("TAM UNG", "tạm ứng"),
    ("TU LAN", "tạm ứng lần"), ("LAN 1", "lần 1"), ("LAN 2", "lần 2"), ("LAN 3", "lần 3"), ("KHAO SAT DIA CHAT", "khảo sát địa chất"),
    ("KHAO SAT", "khảo sát"), ("THIET KE", "thiết kế"), ("GIAM SAT", "giám sát"), ("THUE MAY", "thuê máy"), ("THUE", "thuê"),
    ("DIEN", "điện"), ("NUOC", "nước"), ("VAN PHONG", "văn phòng"), ("VPP", "văn phòng phẩm"), ("XANG DAU", "xăng dầu"),
    ("THAU PHU", "thầu phụ"), ("HOP DONG", "hợp đồng"), ("THANH TOAN", "thanh toán"), ("TIEN", "tiền"), ("HANG", "hàng"),
    ("PHI", "phí"), ("DICH VU", "dịch vụ"), ("SUA CHUA", "sửa chữa"), ("LAP DAT", "lắp đặt"), ("THI CONG", "thi công"),
    ("COP PHA", "cốp pha"), ("GIAN GIAO", "giàn giáo"), ("SON", "sơn"), ("ONG", "ống"), ("GO", "gỗ"), ("DA", "đá"), ("CAU", "cẩu"),
    ("BAO HIEM", "bảo hiểm"), ("THUE GTGT", "thuế GTGT"), ("THUE TNDN", "thuế TNDN"), ("KIEM DINH", "kiểm định"),
]
_TU_DIEN_SORTED = sorted(TU_DIEN, key=lambda x: -len(x[0]))

STOP_NCC = {"CONG", "TY", "CTY", "TNHH", "CO", "PHAN", "CP", "THUONG", "MAI", "TM", "DICH", "VU", "DV", "XAY", "DUNG", "XD", "VA",
            "DAU", "TU", "PHAT", "TRIEN", "SAN", "XUAT", "SX", "TONG", "HOP", "MTV", "MOT", "THANH", "VIEN", "QUOC", "TE", "VIET", "NAM",
            "HA", "NOI", "KY", "THUAT", "KT", "CONG", "NGHE", "CN", "DT", "PT", "NHA", "DO", "THI", "TAP", "DOAN"}


def _phuc_hoi_dau(s: str) -> str:
    out = " " + s.upper() + " "
    for k, v in _TU_DIEN_SORTED:
        out = re.sub(rf"(?<![A-Z]){re.escape(k)}(?![A-Z])", v, out)
    return re.sub(r"\s+", " ", out).strip().lower()


def _key_tokens(name: str) -> list[str]:
    return [t for t in utils.bo_dau(name).replace("-", " ").split() if t.isalpha() and len(t) > 1 and t not in STOP_NCC]


def tim_ncc_gan_dung(ten_tho: str) -> dict | None:
    """Tìm NCC trong danh mục theo các từ khoá chính: 'CTY TIEN PHAT' -> ... 'Kim Tuấn Phát'? (không) ; 'NGOC DAN' -> 'Ngọc Dân'."""
    keys = _key_tokens(ten_tho)
    if not keys:
        return None
    best, best_score = None, 0
    for dt in utils.doi_tuong_list():
        toks = _key_tokens(dt.get("ten", ""))
        if not toks:
            continue
        if all(k in toks for k in keys):
            score = len(keys) / len(toks)  # càng ít từ thừa càng tốt
            if score > best_score:
                best, best_score = dt, score
    return best if best_score >= 0.34 else None


def phan_tich_noi_dung(nd: str, loai: str) -> dict:
    """Trả về dict: muc_dich, cong_trinh (dict|None), ncc_text, doi_tuong (dict|None), dien_giai, do_tin_cay."""
    raw = nd or ""
    s = re.sub(r"-\d{6}-\d{2}:\d{2}:\d{2}\s*\d*\s*$", "", raw.strip())   # bỏ đuôi -090126-08:55:14 520237
    s = re.sub(r"\s+", " ", s)
    u = utils.bo_dau(s)
    muc_dich, ct_text, ncc_text = "", "", ""
    hit = False
    m = re.match(r"^(?:HCIC|CTY|CONG TY)?\s*(?:TT|THANH TOAN|CHUYEN TIEN|TRA|CT)\s+(.*)$", u)
    if loai == "chi" and m:
        hit = True
        rest = m.group(1)
        if " CHO " in rest:
            rest, ncc_text = rest.rsplit(" CHO ", 1)
            ncc_text = re.sub(r"^(CTY|CONG TY|CT|C TY)\s+", "", ncc_text).strip()
        mct = re.search(r"\b(CT|CONG TRINH|DU AN|DA)\s+(.+)$", rest)
        if mct:
            muc_dich = rest[: mct.start()].strip()
            ct_text = mct.group(2).strip()
        elif re.search(r"\bVP\b", rest):
            muc_dich = re.sub(r"\bVP\b", "", rest).strip()
            ct_text = "VP"
        else:
            muc_dich = rest.strip()
    elif loai == "thu":
        # Thu: "<KH> ... TT TIEN ... THEO HD SO ..." -> khách hàng là phần đầu, công trình dò trong toàn chuỗi
        mk = re.search(r"\b(TT|THANH TOAN|CHUYEN TIEN|TRA TIEN)\b", u)
        if mk:
            hit = True
            ncc_text = u[: mk.start()].strip(" -#")
            muc_dich = u[mk.end():].strip()[:80]
            ncc_text = re.sub(r"^CONG TY CP XAY DUNG DAN DUNG CONG NGHIEP HA NOI#*\d*\s*", "", ncc_text)  # bỏ tên công ty mình nếu NH ghi đầu
    cong_trinh = utils.match_cong_trinh(ct_text) if ct_text else utils.match_cong_trinh(s)
    doi_tuong = utils.match_doi_tuong(s) or (tim_ncc_gan_dung(ncc_text) if ncc_text else None)
    ten_dt = doi_tuong["ten"] if doi_tuong else (ncc_text.title() if ncc_text else "")
    if hit:
        verb = "Thanh toán" if loai == "chi" else "Thu tiền"
        parts = [verb, _phuc_hoi_dau(muc_dich) if muc_dich else ""]
        if cong_trinh and cong_trinh["ma"] != "VP":
            parts.append("công trình " + cong_trinh["ten"].replace("Thi công công trình ", "").replace("Dự án thi công công trình ", "")[:50])
        elif ct_text == "VP":
            parts.append("văn phòng")
        if ten_dt:
            parts.append(("cho " if loai == "chi" else "từ ") + ten_dt)
        dien_giai = re.sub(r"\s+", " ", " ".join(p for p in parts if p)).strip()
        dien_giai = re.sub(r"\b(tiền|thanh toán|thu tiền) \1\b", r"\1", dien_giai)
        dien_giai = re.sub(r"\bThu tiền tiền\b", "Thu tiền", dien_giai)
        dien_giai = dien_giai[0].upper() + dien_giai[1:]
    else:
        dien_giai = s[:200]
    tin_cay = 0.4
    if hit and muc_dich:
        tin_cay = 0.6
    if hit and (cong_trinh or ct_text == "VP") and (doi_tuong or ncc_text):
        tin_cay = 0.75
    if doi_tuong:
        tin_cay = max(tin_cay, 0.7)
    return {"muc_dich": muc_dich, "cong_trinh": cong_trinh, "ncc_text": ncc_text, "doi_tuong": doi_tuong, "ten_doi_tuong": ten_dt,
            "dien_giai": dien_giai, "do_tin_cay": tin_cay}
