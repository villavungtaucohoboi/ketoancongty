"""Ghi file Excel nhập khẩu MISA SME.NET cho phân hệ Mua hàng (hóa đơn mua hàng hóa/dịch vụ)."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .. import utils


def _ten(lst: list[dict], ma: str) -> str:
    for x in lst:
        if x["ma"] == ma:
            return x["ten"]
    return ma or ""


def build_rows(items: list[dict]) -> list[list]:
    """items: [{'hd': HoaDon dict, 'pl': PhanLoai dict}]"""
    cfg = utils.settings()["hoa_don"]
    rows = []
    for i, it in enumerate(items, start=1):
        hd, pl = it["hd"], it["pl"]
        dt = utils.match_doi_tuong(hd.get("ten_nguoi_ban", ""), hd.get("mst_nguoi_ban", ""))
        ma_ncc = dt["ma"] if dt else f"NCC_{(hd.get('mst_nguoi_ban') or '')[-6:]}"
        ghi_chu_ncc = "" if dt else "NCC chưa có trong danh mục - thêm vào MISA trước khi nhập"
        ngay = hd.get("ngay_hoa_don", "")
        d = utils.parse_date(ngay)
        so_ct = f"{cfg['so_ct_mua_hang']}{d:%y%m%d}-{i:03d}" if d else f"{cfg['so_ct_mua_hang']}-{i:03d}"
        tin_cay = min(float(hd.get("do_tin_cay", 1)), float(pl.get("do_tin_cay", 1)))
        ghi_chu = f"Nguồn {hd.get('nguon','')}; tin cậy {tin_cay:.0%}" + (" - CẦN KIỂM TRA" if tin_cay < 0.7 else "") + (("; " + ghi_chu_ncc) if ghi_chu_ncc else "")
        dong_hang = hd.get("dong_hang") or [{"ten_hang": hd.get("ghi_chu") or "Hàng hóa, dịch vụ theo HĐ", "dvt": "", "so_luong": 1,
                                               "don_gia": hd.get("tien_hang", 0), "thanh_tien": hd.get("tien_hang", 0),
                                               "thue_suat": None, "tien_thue": hd.get("tien_thue", 0)}]
        for dh in dong_hang:
            rows.append([
                ngay, ngay, so_ct, ma_ncc, hd.get("ten_nguoi_ban", ""), hd.get("mst_nguoi_ban", ""), hd.get("dia_chi_nguoi_ban", ""),
                f"Mua {dh.get('ten_hang','')} của {hd.get('ten_nguoi_ban','')} theo HĐ {hd.get('so_hoa_don','')}"[:250],
                hd.get("mau_so", ""), hd.get("ky_hieu", ""), hd.get("so_hoa_don", ""), ngay,
                "", dh.get("ten_hang", ""), dh.get("dvt", ""), dh.get("so_luong", 0), dh.get("don_gia", 0), dh.get("thanh_tien", 0),
                dh.get("thue_suat") if dh.get("thue_suat") is not None else "KCT", dh.get("tien_thue", 0),
                cfg["tk_no_mac_dinh"], cfg["tk_co_mac_dinh"], cfg["tk_thue_gtgt"],
                _ten(utils.khoan_muc_list(), pl.get("khoan_muc_ma", "")), _ten(utils.cong_trinh_list(), pl.get("cong_trinh_ma", "")),
                hd.get("ma_cqt", ""), ghi_chu,
            ])
    return rows


def xuat_file_mua_hang(items: list[dict], out_path: str | Path) -> Path:
    tpl = utils.misa_templates()["mua_hang"]
    wb = Workbook()
    ws = wb.active
    ws.title = tpl["sheet"]
    ws.append(tpl["cot"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in build_rows(items):
        ws.append(r)
        if "CẦN KIỂM TRA" in str(r[-1]) or "chưa có trong danh mục" in str(r[-1]):
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", fgColor="FFF2CC")
    widths = [12, 12, 14, 14, 40, 13, 30, 50, 8, 10, 10, 12, 8, 40, 8, 10, 14, 14, 8, 14, 8, 8, 8, 28, 28, 36, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col in (16, 17, 18, 20):
        for cell in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            cell[0].number_format = "#,##0"
    ws.freeze_panes = "A2"
    out_path = Path(out_path)
    utils.ensure_dir(out_path.parent)
    wb.save(out_path)
    return out_path
