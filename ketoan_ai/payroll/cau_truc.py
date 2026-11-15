"""Nhận diện cấu trúc sheet "Bảng lương" của công ty và bộ CÔNG THỨC CHUẨN.

- nhan_dien(ws): tự tìm dòng tiêu đề khối (Khối văn phòng / Thực tập sinh / Khối thi công), các dòng
  công trình, dòng Tổng cộng và dòng nhân viên thuộc khối nào. Không phụ thuộc số dòng cố định.
- trich_cong_thuc(file_chuan): đọc file bảng lương CHUẨN (có đủ công thức), lưu công thức theo dạng mẫu
  ({r} = dòng của người đó, {hdr_vp} / {hdr_tc} = dòng tiêu đề khối) vào config/luong_cong_thuc.yaml.
  Công thức riêng của từng người (vd. làm thêm 70.000đ/buổi) được giữ theo họ tên.
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field

import openpyxl
from openpyxl.utils import get_column_letter


def norm(s) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(s or ""))).strip().casefold()


@dataclass
class CauTruc:
    hdr: dict[str, int] = field(default_factory=dict)        # vp / tts / tc -> dòng tiêu đề khối
    nhom: list[tuple[str, int, int, int]] = field(default_factory=list)  # (loại, dòng nhãn, dòng đầu, dòng cuối)
    nhan_vien: dict[int, tuple[str, str]] = field(default_factory=dict)  # dòng -> (khối, tên chuẩn hóa)
    dong_tong: int | None = None


def _gia_tri(v):
    return getattr(v, "text", v)


def nhan_dien(ws) -> CauTruc:
    ct = CauTruc()
    khoi = None
    nhan = []  # (loại, dòng)
    for r in range(1, ws.max_row + 1):
        a, b = ws[f"A{r}"].value, ws[f"B{r}"].value
        na = norm(a)
        if isinstance(a, (int, float)) and b:
            if khoi:
                ct.nhan_vien[r] = (khoi, norm(b))
            continue
        if na.startswith("khối văn phòng"):
            khoi, ct.hdr["vp"] = "vp", r
            nhan.append(("vp", r))
        elif na.startswith("thực tập sinh"):
            khoi, ct.hdr["tts"] = "tts", r
            nhan.append(("tts", r))
        elif na.startswith("khối thi công"):
            khoi, ct.hdr["tc"] = "tc", r
            nhan.append(("tc", r))
        elif na.startswith("công trình") and khoi == "tc":
            nhan.append(("ct", r))
        elif norm(b) == "tổng cộng" or na == "tổng cộng":
            ct.dong_tong = r
            nhan.append(("tong", r))
    # phạm vi dòng nhân viên của từng nhãn
    nv = sorted(ct.nhan_vien)
    for i, (loai, r) in enumerate(nhan):
        if loai == "tong":
            ds = [x for x in nv if x < r]
        elif loai == "tc":
            nxt = next((rr for ll, rr in nhan[i + 1:] if ll in ("tong",)), 10 ** 6)
            ds = [x for x in nv if r < x < nxt]
        else:
            nxt = next((rr for ll, rr in nhan[i + 1:]), 10 ** 6)
            ds = [x for x in nv if r < x < nxt]
        if ds:
            ct.nhom.append((loai, r, ds[0], ds[-1]))
    return ct


# ---------------------------------------------------------------------------
_REF = re.compile(r"(?<![A-Za-z_$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")


def _mau_hoa(f: str, r: int, hdr: dict[str, int]) -> str:
    """Đổi công thức của dòng r thành mẫu: số dòng r -> {r}; dòng tiêu đề tuyệt đối -> {hdr_x}."""
    rev = {v: k for k, v in hdr.items()}

    def thay(m):
        d1, col, d2, row = m.group(1), m.group(2), m.group(3), int(m.group(4))
        if not d2 and row == r:
            return f"{d1}{col}{{r}}"
        if d2 and row in rev:
            return f"{d1}{col}${{hdr_{rev[row]}}}"
        if not d2 and row in rev:
            return f"{d1}{col}{{hdr_{rev[row]}}}"
        return m.group(0)

    # không đụng phần trong chuỗi '...'
    parts = re.split(r"('[^']*')", f)
    return "".join(p if p.startswith("'") else _REF.sub(thay, p) for p in parts)


def ap_mau(mau: str, r: int, hdr: dict[str, int]) -> str:
    return mau.replace("{r}", str(r)).replace("{hdr_vp}", str(hdr.get("vp", 7))).replace(
        "{hdr_tts}", str(hdr.get("tts", 0))).replace("{hdr_tc}", str(hdr.get("tc", 0)))


def trich_cong_thuc(file_chuan) -> dict:
    wb = openpyxl.load_workbook(file_chuan)
    ws = wb["Bảng lương"] if "Bảng lương" in wb.sheetnames else wb[[s for s in wb.sheetnames if norm(s) == norm("Bảng lương")][0]]
    ct = nhan_dien(ws)
    hdr = ct.hdr
    cols = [get_column_letter(c) for c in range(3, ws.max_column + 1)]
    theo_nguoi: dict[str, dict[str, str]] = {}
    theo_khoi: dict[str, list[dict]] = {}
    for r, (khoi, ten) in ct.nhan_vien.items():
        d = {}
        for c in cols:
            v = _gia_tri(ws[f"{c}{r}"].value)
            if isinstance(v, str) and v.startswith("="):
                d[c] = _mau_hoa(v, r, hdr)
        theo_nguoi[ten] = d
        theo_khoi.setdefault(khoi, []).append(d)
    mac_dinh, rieng = {}, {}
    for khoi, ds in theo_khoi.items():
        md = {}
        for c in cols:
            cnt = Counter(d.get(c) for d in ds)
            best, _ = cnt.most_common(1)[0]
            if best:
                md[c] = best
        mac_dinh[khoi] = md
    for r, (khoi, ten) in ct.nhan_vien.items():
        d, md = theo_nguoi[ten], mac_dinh[khoi]
        khac = {c: d.get(c) for c in sorted(set(d) | set(md)) if d.get(c) != md.get(c)}
        if khac:
            rieng[ten] = khac  # None = người này không dùng công thức ở cột đó
    # hằng số / công thức ở dòng tiêu đề khối và dòng nhóm
    hang = {}
    for loai, r in [(k, v) for k, v in hdr.items()]:
        hang[loai] = {c: (_mau_hoa(v, r, hdr) if isinstance(v, str) and v.startswith("=") else v)
                      for c in cols for v in [_gia_tri(ws[f"{c}{r}"].value)]
                      if v not in (None, "") and c not in ("A", "B", "E")
                      and not (isinstance(v, str) and v.upper().startswith("=SUBTOTAL"))}  # tổng nhóm dựng lại theo vùng dòng mới
    nhom = {}
    for loai, r, a, b in ct.nhom:
        dd = {}
        for c in cols:
            v = _gia_tri(ws[f"{c}{r}"].value)
            if isinstance(v, str) and v.upper().startswith("=SUBTOTAL(9,"):
                dd[c] = "SUBTOTAL"
        nhom.setdefault(loai, dd)
    return {"mac_dinh": mac_dinh, "rieng": rieng, "hang_so": hang, "nhom_subtotal": nhom}
