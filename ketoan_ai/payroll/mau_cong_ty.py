"""Cập nhật bảng lương tháng mới trên chính FILE MẪU của công ty, không đổi công thức.

Luồng: file lương tháng trước (nhân sự đã cập nhật sheet "Chấm công + đánh giá hiệu quả")
  -> chép ngày công / hệ số / làm thêm sang các ô nhập tay của sheet "Bảng lương" (khớp theo họ tên)
  -> đổi tiêu đề tháng, chuyển số tháng trước sang bảng tổng hợp
  -> ghi trực tiếp vào XML của file (giữ nguyên định dạng, logo, chú thích, sheet ẩn, mọi công thức)
  -> tính thử bằng LibreOffice (nếu máy có) để xem trước + kiểm tra; Excel luôn tự tính lại khi mở file.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter

from .. import utils


def _norm(s) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(s or ""))).strip().casefold()


def _ten_sheet_that(ten_co: list[str], ten: str) -> str:
    """Trả về tên sheet thật trong file khớp với tên cấu hình (bỏ qua khác biệt mã Unicode, hoa/thường, khoảng trắng)."""
    k = _norm(ten)
    for t in ten_co:
        if _norm(t) == k:
            return t
    raise ValueError(f"Không tìm thấy sheet '{ten}' trong file. Các sheet có trong file: {', '.join(ten_co)}. "
                     "Kiểm tra đúng file bảng lương của công ty chưa.")


def _so(v) -> float:
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return utils.parse_money(v)


@dataclass
class KetQua:
    thay_doi: pd.DataFrame
    canh_bao: list[str] = field(default_factory=list)
    bang_luong: pd.DataFrame | None = None
    tong: dict[str, float] = field(default_factory=dict)
    da_tinh: bool = False          # True nếu đã tính thử bằng LibreOffice
    thang_cu: str = ""
    thang_moi: str = ""


# ---------------------------------------------------------------------------
# Ghi ô trực tiếp vào XML của sheet (giữ nguyên style)
# ---------------------------------------------------------------------------
_CELL_RE = r'<c r="{ref}"(?P<attrs>[^>/]*)(?:/>|>(?P<inner>.*?)</c>)'


def _col_num(ref: str) -> int:
    return column_index_from_string(re.match(r"[A-Z]+", ref).group(0))


def _cell_xml(ref: str, attrs: str, value) -> str:
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    if value is None:
        return f'<c r="{ref}"{attrs}/>'
    if isinstance(value, str):
        return f'<c r="{ref}"{attrs} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    return f'<c r="{ref}"{attrs}><v>{repr(float(value)) if not float(value).is_integer() else int(value)}</v></c>'


def _ghi_o(xml: str, ref: str, value) -> str:
    m = re.search(_CELL_RE.format(ref=ref), xml, re.S)
    if m:
        if m.group("inner") and "<f" in m.group("inner"):
            raise ValueError(f"Ô {ref} là công thức, không được ghi đè.")
        return xml[: m.start()] + _cell_xml(ref, m.group("attrs"), value) + xml[m.end():]
    row = int(re.search(r"\d+", ref).group(0))
    mr = re.search(rf'<row r="{row}"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S)
    if not mr:
        raise ValueError(f"Không thấy dòng {row} trong sheet để ghi ô {ref}.")
    new = _cell_xml(ref, "", value)
    if mr.group(0).endswith("/>"):
        rowtag = mr.group(0)[:-2] + ">"
        return xml[: mr.start()] + rowtag + new + "</row>" + xml[mr.end():]
    inner, start_inner = mr.group(1), mr.start(1)
    pos = len(inner)
    for mc in re.finditer(r'<c r="([A-Z]+)\d+"', inner):
        if _col_num(mc.group(1)) > _col_num(ref):
            pos = mc.start()
            break
    return xml[: start_inner + pos] + new + xml[start_inner + pos:]


def _sheet_parts(z: zipfile.ZipFile) -> dict[str, str]:
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    out = {}
    for m in re.finditer(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wbx):
        t = re.search(rf'<Relationship\b[^>]*?Id="{m.group(2)}"[^>]*?Target="([^"]+)"', rels) or \
            re.search(rf'<Relationship\b[^>]*?Target="([^"]+)"[^>]*?Id="{m.group(2)}"', rels)
        target = t.group(1).lstrip("/")
        out[_unescape(m.group(1))] = target if target.startswith("xl/") else "xl/" + target
    return out


def _unescape(s: str) -> str:
    return s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'")


def _viet_lai_zip(src: Path, dst: Path, sua: dict[str, bytes]):
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = sua.get(item.filename, zi.read(item.filename))
            if data is None:
                continue  # bỏ phần này khỏi file
            zo.writestr(item, data)


# ---------------------------------------------------------------------------
# Tính thử bằng LibreOffice (bản nháp: XLOOKUP -> INDEX/MATCH chỉ trong bản nháp)
# ---------------------------------------------------------------------------
def _tach_doi_so(s: str) -> tuple[list[str], int]:
    """s bắt đầu ngay sau 'XLOOKUP(' ; trả về (các đối số, vị trí dấu ')' đóng)."""
    args, depth, cur, i, in_str = [], 0, "", 0, False
    while i < len(s):
        ch = s[i]
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            j = s.index("'", i + 1)
            cur += s[i: j + 1]
            i = j + 1
            continue
        if not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                if depth == 0:
                    args.append(cur)
                    return args, i
                depth -= 1
            elif ch == "," and depth == 0:
                args.append(cur)
                cur, i = "", i + 1
                continue
        cur += ch
        i += 1
    raise ValueError("Công thức XLOOKUP không đóng ngoặc.")


def doi_xlookup(f: str) -> str:
    key = "_xlfn.XLOOKUP("
    while key in f:
        k = f.rindex(key)  # xử lý trong cùng trước
        args, end = _tach_doi_so(f[k + len(key):])
        a = args + [""] * (4 - len(args))
        repl = f"INDEX({a[2]},MATCH({a[0]},{a[1]},0))"
        if a[3].strip():
            repl = f"IFERROR({repl},{a[3]})"
        f = f[:k] + repl + f[k + len(key) + end + 1:]
    return f


MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="KTRecalc" script:language="StarBasic">
Sub RecalcSave()
  ThisComponent.calculateAll()
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
</script:module>"""


def _soffice() -> str | None:
    for c in ("soffice", "libreoffice", r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if shutil.which(c) or Path(c).exists():
            return shutil.which(c) or c
    return None


def tinh_thu(path: Path, timeout: int = 180) -> Path | None:
    """Tạo bản nháp (XLOOKUP đổi sang INDEX/MATCH), cho LibreOffice tính lại, trả về đường dẫn bản nháp đã tính."""
    so = _soffice()
    if not so:
        return None
    work = Path(tempfile.mkdtemp(prefix="kt-luong-"))
    nhap = work / "nhap.xlsx"
    with zipfile.ZipFile(path) as z:
        sua = {}
        for n in z.namelist():
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                x = z.read(n).decode("utf-8")
                if "_xlfn.XLOOKUP" in x:
                    x = re.sub(r"(<f\b[^>]*?(?<!/)>)(.*?)(</f>)", lambda m: m.group(1) + escape(doi_xlookup(_unescape(m.group(2)))) + m.group(3), x, flags=re.S)
                    sua[n] = x.encode("utf-8")
    _viet_lai_zip(path, nhap, sua)
    prof = work / "profile"
    env = dict(os.environ, HOME=str(work))
    subprocess.run([so, "--headless", "--terminate_after_init", f"-env:UserInstallation={prof.as_uri()}"],
                   capture_output=True, timeout=timeout, env=env)
    mdir = prof / "user" / "basic" / "Standard"
    if not mdir.exists():
        return None
    (mdir / "KTRecalc.xba").write_text(MACRO, encoding="utf-8")
    lst = (mdir / "script.xlb").read_text(encoding="utf-8")
    if "KTRecalc" not in lst:
        (mdir / "script.xlb").write_text(lst.replace("</library:library>", '<library:element library:name="KTRecalc"/></library:library>'), encoding="utf-8")
    truoc = nhap.stat().st_mtime_ns
    subprocess.run([so, "--headless", "--norestore", f"-env:UserInstallation={prof.as_uri()}",
                    "vnd.sun.star.script:Standard.KTRecalc.RecalcSave?language=Basic&location=application", str(nhap)],
                   capture_output=True, timeout=timeout, env=env)
    return nhap if nhap.stat().st_mtime_ns != truoc else None


def _gan_gia_tri_cache(out: Path, da_tinh: Path):
    """Đưa giá trị LibreOffice đã tính vào file kết quả (công thức giữ nguyên) để xem trước đúng số ngay."""
    wv = openpyxl.load_workbook(da_tinh, data_only=True, read_only=True)
    vals = {ws.title: {(c.row, c.column): c.value for row in ws.iter_rows() for c in row if getattr(c, "row", None)} for ws in wv.worksheets}
    with zipfile.ZipFile(out) as z:
        parts = _sheet_parts(z)
        sua = {}
        vals_n = {_norm(k): v for k, v in vals.items()}
        for name, part in parts.items():
            if _norm(name) not in vals_n:
                continue
            x = z.read(part).decode("utf-8")
            v = vals_n[_norm(name)]

            def thay(m):
                ref, attrs, inner = m.group(1), m.group(2), m.group(3)
                if "<f" not in inner:
                    return m.group(0)
                col = column_index_from_string(re.match(r"[A-Z]+", ref).group(0))
                row = int(re.search(r"\d+", ref).group(0))
                val = v.get((row, col))
                fpart = re.search(r"<f\b.*?(?:/>|</f>)", inner, re.S).group(0)
                attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
                if val is None:
                    return f'<c r="{ref}"{attrs}>{fpart}</c>'
                if isinstance(val, bool):
                    return f'<c r="{ref}"{attrs} t="b">{fpart}<v>{int(val)}</v></c>'
                if isinstance(val, (int, float)):
                    return f'<c r="{ref}"{attrs}>{fpart}<v>{val}</v></c>'
                sval = str(val)
                if sval.startswith("#"):
                    return m.group(0)  # ô lỗi sẵn trong file gốc (#REF!...): giữ nguyên, Excel tự tính khi mở
                return f'<c r="{ref}"{attrs} t="str">{fpart}<v>{escape(sval)}</v></c>'

            sua[part] = re.sub(r'<c r="([A-Z]+[0-9]+)"([^>/]*)>(.*?)</c>', thay, x, flags=re.S).encode("utf-8")
    tmp = out.with_suffix(".tmp.xlsx")
    _viet_lai_zip(out, tmp, sua)
    tmp.replace(out)


# ---------------------------------------------------------------------------
# Hàm chính
# ---------------------------------------------------------------------------
def _thang_trong_tieu_de(s: str) -> str:
    m = re.search(r"TH[ÁA]NG\s*(\d{1,2})\s*[/.]\s*(\d{4})", unicodedata.normalize("NFC", str(s or "")), re.I)
    return f"{int(m.group(1)):02d}/{m.group(2)}" if m else ""


def _serial(d: date) -> int:
    return (d - date(1899, 12, 30)).days


def _la_cong_thuc(v) -> bool:
    return (isinstance(v, str) and v.startswith("=")) or hasattr(v, "text")


def _trong(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _ghi_cong_thuc(xml: str, ref: str, f: str) -> str:
    """Ghi công thức vào ô ĐANG TRỐNG (giữ style)."""
    body = escape(f[1:] if f.startswith("=") else f)
    m = re.search(_CELL_RE.format(ref=ref), xml, re.S)
    if m:
        if m.group("inner") and ("<f" in m.group("inner") or "<v" in m.group("inner") or "<is" in m.group("inner")):
            raise ValueError(f"Ô {ref} không trống, không ghi đè.")
        attrs = re.sub(r'\s+t="[^"]*"', "", m.group("attrs"))
        return xml[: m.start()] + f'<c r="{ref}"{attrs}><f>{body}</f></c>' + xml[m.end():]
    x = _ghi_o(xml, ref, 0)  # tạo ô đúng vị trí rồi thay nội dung
    return re.sub(rf'<c r="{ref}"([^>/]*)><v>0</v></c>', rf'<c r="{ref}"\1><f>{body}</f></c>', x, count=1)


def _o_tham_chieu(f) -> str | None:
    """'='Bảng lương'!AD40' -> 'AD40'"""
    m = re.search(r"!\$?([A-Z]{1,3})\$?(\d+)\s*$", str(getattr(f, "text", f) or ""))
    return f"{m.group(1)}{m.group(2)}" if m else None


def cap_nhat(file_vao: str | Path, thang_moi: str, file_ra: str | Path, tinh: bool = True) -> KetQua:
    """thang_moi: 'mm/yyyy'. Trả về KetQua; file_ra là bảng lương tháng mới (công thức có sẵn giữ nguyên)."""
    from . import cau_truc

    cfg = utils.load_yaml("luong_hcic.yaml")
    chuan = utils.load_yaml("luong_cong_thuc.yaml")
    file_vao, file_ra = Path(file_vao), Path(file_ra)
    mm, yyyy = [int(x) for x in thang_moi.split("/")]
    wf = openpyxl.load_workbook(file_vao)                 # công thức
    wv = openpyxl.load_workbook(file_vao, data_only=True)  # giá trị đang lưu
    that = lambda ten: _ten_sheet_that(wf.sheetnames, ten)  # noqa: E731
    SH_BL, SH_CC, SH_TH = that(cfg["sheet_bang_luong"]), that(cfg["sheet_cham_cong"]), that(cfg["sheet_tong_hop"])
    try:
        SH_PC = that(cfg["sheet_phu_cap"])
    except ValueError:
        SH_PC = None
    tieu_de = [[that(sh), o] for sh, o in cfg["tieu_de"]]
    bl_f, bl_v, cc = wf[SH_BL], wv[SH_BL], wv[SH_CC]
    ct = cau_truc.nhan_dien(bl_f)
    if not ct.nhan_vien or "vp" not in ct.hdr:
        raise ValueError("Không nhận ra cấu trúc sheet Bảng lương (thiếu nhãn 'Khối văn phòng' / 'Khối thi công' ở cột A).")
    canh_bao: list[str] = []
    TEN_KHOI = {"vp": "Văn phòng", "tts": "Thực tập sinh", "tc": "Thi công"}

    # --- chấm công & phụ cấp theo tên
    def theo_ten(ws, r0, cot_ten="B"):
        d = {}
        for r in range(r0, ws.max_row + 1):
            ten = ws[f"{cot_ten}{r}"].value
            if ten and isinstance(ws[f"A{r}"].value, (int, float)) or (ten and ws is not cc and r >= r0 and ws[f"A{r}"].value not in (None, "")):
                k = _norm(ten)
                if k in d and ws is cc:
                    canh_bao.append(f"Tên '{ten}' xuất hiện 2 lần trong sheet chấm công (dòng {d[k]} và {r}); dùng dòng đầu.")
                d.setdefault(k, r)
        return d

    cc_rows = theo_ten(cc, cfg["dong_bat_dau_cham_cong"])
    pc = wv[SH_PC] if SH_PC else None
    pc_rows = {}
    if pc is not None:
        for r in range(2, pc.max_row + 1):
            if pc[f"B{r}"].value:
                pc_rows.setdefault(_norm(pc[f"B{r}"].value), r)

    edits: dict[str, dict[str, object]] = {SH_BL: {}, SH_TH: {}}
    cong_thuc: dict[str, str] = {}   # ô -> công thức cần điền (chỉ ô trống)
    thay_doi = []
    hdr = ct.hdr

    def ghi_cong_thuc_neu_trong(o, mau, r):
        if mau and _trong(bl_f[o].value):
            cong_thuc[o] = cau_truc.ap_mau(mau, r, hdr)

    # 1) dòng tiêu đề khối: hằng số/công thức chuẩn nếu trống
    for loai, r in hdr.items():
        for c, v in (chuan.get("hang_so", {}).get(loai) or {}).items():
            o = f"{c}{r}"
            if not _trong(bl_f[o].value):
                continue
            if isinstance(v, str) and v.startswith("="):
                cong_thuc[o] = cau_truc.ap_mau(v, r, hdr)
            else:
                edits[SH_BL][o] = v
    # 2) dòng nhóm (khối, công trình, tổng): SUBTOTAL theo vùng dòng mới nếu trống
    for loai, r, a, b in ct.nhom:
        for c in (chuan.get("nhom_subtotal", {}).get(loai) or {}):
            if _trong(bl_f[f"{c}{r}"].value) and not (loai in hdr and c in (chuan.get("hang_so", {}).get(loai) or {})):
                cong_thuc[f"{c}{r}"] = f"=SUBTOTAL(9,{c}{a}:{c}{b})"
    # 3) từng nhân viên
    so_cong_thuc_nv = 0
    for r, (khoi, k) in sorted(ct.nhan_vien.items()):
        ten = bl_v[f"B{r}"].value
        rieng = (chuan.get("rieng") or {}).get(k, {})
        mac_dinh = (chuan.get("mac_dinh") or {}).get(khoi, {})
        for c in set(mac_dinh) | set(rieng):
            mau = rieng[c] if c in rieng else mac_dinh.get(c)
            truoc = len(cong_thuc)
            ghi_cong_thuc_neu_trong(f"{c}{r}", mau, r)
            so_cong_thuc_nv += len(cong_thuc) - truoc
        if pc is not None and k in pc_rows:
            for c_bl, c_pc in cfg["tu_phu_cap"].items():
                o = f"{c_bl}{r}"
                if _trong(bl_f[o].value) and o not in cong_thuc:
                    v = pc[f"{c_pc}{pc_rows[k]}"].value
                    if isinstance(v, (int, float)) and v:
                        edits[SH_BL][o] = v
        if k not in cc_rows:
            canh_bao.append(f"'{ten}' (dòng {r} Bảng lương) không có trong sheet chấm công: giữ nguyên số liệu, cần kiểm tra (nghỉ việc?).")
            continue
        rc = cc_rows[k]
        for c_bl, c_cc in cfg["anh_xa"][khoi].items():
            o = f"{c_bl}{r}"
            if _la_cong_thuc(bl_f[o].value) or o in cong_thuc:
                continue  # ô công thức (vd. công TTS lấy bằng XLOOKUP): không đụng
            cu, moi = _so(bl_v[o].value), _so(cc[f"{c_cc}{rc}"].value)
            if abs(cu - moi) > 1e-9:
                thay_doi.append({"Dòng": r, "Họ tên": ten, "Khối": TEN_KHOI[khoi], "Ô": o,
                                 "Nội dung": str(bl_v[f"{c_bl}5"].value or bl_v[f"{c_bl}6"].value or c_bl).replace("\n", " "),
                                 "Cũ": cu, "Mới": moi})
                edits[SH_BL][o] = moi if moi else (None if bl_v[o].value is None else 0)
    da_co = {k for _, k in ct.nhan_vien.values()}
    for k, r in cc_rows.items():
        if k not in da_co:
            canh_bao.append(f"'{cc[f'B{r}'].value}' có trong chấm công nhưng chưa có dòng trong Bảng lương: "
                            "thêm dòng vào đúng khối bằng Excel (chép dòng đồng nghiệp cùng khối) rồi chạy lại.")
    if so_cong_thuc_nv:
        canh_bao.insert(0, f"File tải lên bị thiếu {so_cong_thuc_nv} công thức ở sheet Bảng lương; app đã điền lại theo công thức chuẩn "
                           "của công ty (kể cả công thức riêng từng người). Ô đang có công thức hoặc số liệu không bị thay đổi.")

    # --- tiêu đề tháng + bảng tổng hợp
    thang_cu = ""
    for sh, o in tieu_de:
        txt = wv[sh][o].value
        t = _thang_trong_tieu_de(txt)
        thang_cu = thang_cu or t
        if txt and t:
            edits.setdefault(sh, {})[o] = re.sub(r"(TH[ÁA]NG\s*)\d{1,2}\s*[/.]\s*\d{4}", rf"\g<1>{mm:02d}/{yyyy}", unicodedata.normalize("NFC", txt), flags=re.I)
    th = cfg["tong_hop"]
    th_f, th_v = wf[SH_TH], wv[SH_TH]
    nguon = {d: _o_tham_chieu(th_f[f"{th['cot_nay']}{d}"].value) for d in th["dong"]}
    if thang_cu and thang_cu != f"{mm:02d}/{yyyy}":
        edits[SH_TH][th["o_thang_nay"]] = _serial(date(yyyy, mm, 1))
        om, oy = [int(x) for x in thang_cu.split("/")]
        edits[SH_TH][th["o_thang_truoc"]] = _serial(date(oy, om, 1))
        for d, o_bl in nguon.items():
            cu = _so(bl_v[o_bl].value) if o_bl else 0
            if cu:
                edits[SH_TH][f"{th['cot_truoc']}{d}"] = cu
            else:
                canh_bao.append(f"Bảng tổng hợp dòng {d}: file tải lên chưa có số thực nhận tháng {thang_cu} (đang 0) nên cột tháng trước "
                                "được giữ nguyên; nhập tay số tháng trước nếu cần so sánh.")

    # --- ghi XML
    file_ra.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(file_vao) as z:
        parts = _sheet_parts(z)
        sua = {}
        for sh in set(edits) | ({SH_BL} if cong_thuc else set()):
            cells = edits.get(sh, {})
            part = parts[_ten_sheet_that(list(parts), sh)]
            x = z.read(part).decode("utf-8")
            for o, v in cells.items():
                x = _ghi_o(x, o, v)
            if sh == SH_BL:
                for o, f in cong_thuc.items():
                    x = _ghi_cong_thuc(x, o, f)
            sua[part] = x.encode("utf-8")
        wbx = z.read("xl/workbook.xml").decode("utf-8")
        if "fullCalcOnLoad" not in wbx:
            wbx = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wbx, count=1) if "<calcPr" in wbx else \
                wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        sua["xl/workbook.xml"] = wbx.encode("utf-8")
        # chuỗi tính (calcChain) cũ không còn khớp khi thêm công thức -> bỏ, Excel tự dựng lại
        if "xl/calcChain.xml" in z.namelist():
            sua["xl/calcChain.xml"] = None
            sua["[Content_Types].xml"] = re.sub(r'<Override[^>]*calcChain[^>]*/>', "", z.read("[Content_Types].xml").decode("utf-8")).encode("utf-8")
            sua["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", z.read("xl/_rels/workbook.xml.rels").decode("utf-8")).encode("utf-8")
    _viet_lai_zip(file_vao, file_ra, sua)

    kq = KetQua(thay_doi=pd.DataFrame(thay_doi, columns=["Dòng", "Họ tên", "Khối", "Ô", "Nội dung", "Cũ", "Mới"]),
                canh_bao=canh_bao, thang_cu=thang_cu, thang_moi=f"{mm:02d}/{yyyy}")

    # --- tính thử để xem trước
    if tinh:
        try:
            nhap = tinh_thu(file_ra)
        except Exception as ex:  # noqa: BLE001
            nhap = None
            canh_bao.append(f"Không tính thử được bằng LibreOffice ({ex}); mở file bằng Excel, Excel sẽ tự tính.")
        if nhap:
            _gan_gia_tri_cache(file_ra, nhap)
            kq.da_tinh = True
            shutil.rmtree(nhap.parent, ignore_errors=True)
    wr = openpyxl.load_workbook(file_ra, data_only=True)
    bl = wr[_ten_sheet_that(wr.sheetnames, SH_BL)]
    rows = []
    for r, (khoi, _) in sorted(ct.nhan_vien.items()):
        rows.append({"STT": bl[f"A{r}"].value, "Họ tên": bl[f"B{r}"].value, "Khối": TEN_KHOI[khoi], "Phòng ban": bl[f"D{r}"].value,
                     **{k: bl[f"{c}{r}"].value for k, c in cfg["cot_ket_qua"].items()}})
    kq.bang_luong = pd.DataFrame(rows)
    for r, (khoi, _) in sorted(ct.nhan_vien.items()):
        if not _so(bl[f"G{r}"].value):
            canh_bao.append(f"'{bl[f'B{r}'].value}' (dòng {r}) có Thu nhập thỏa thuận = 0 nên chỉ được tính phụ cấp/tiền ăn: "
                            "kiểm tra lại cột G trong sheet Bảng lương.")
    nhan = {d: str(th_v[f"B{d}"].value or f"Dòng {d}") for d in th["dong"]}
    kq.tong = {nhan[d]: _so(bl[o].value) for d, o in nguon.items() if o}
    return kq
