"""Ghi file Excel nhập khẩu MISA SME.NET cho phân hệ Ngân hàng (Thu tiền gửi / Chi tiền gửi)."""
from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .. import utils
from ..ai import GiaoDichNganHang

FILL_LOW = PatternFill("solid", fgColor="FFF2CC")   # vàng nhạt: AI chưa chắc, cần kiểm tra
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")


def _ten_cong_trinh(ma: str) -> str:
    for c in utils.cong_trinh_list():
        if c["ma"] == ma:
            return c["ten"]
    return ma or ""


def _ten_khoan_muc(ma: str) -> str:
    for k in utils.khoan_muc_list():
        if k["ma"] == ma:
            return k["ten"]
    return ma or ""


def build_rows(giao_dich: list[dict], phan_loai: list[GiaoDichNganHang], tab_cfg: dict) -> tuple[list[list], list[list]]:
    """Trả về (rows_thu, rows_chi) theo đúng thứ tự cột trong misa_templates.yaml."""
    cfg = utils.settings()["sao_ke"]
    rows_thu, rows_chi = [], []
    counters = {"thu": 0, "chi": 0}
    for gd, pl in zip(giao_dich, phan_loai):
        loai = pl.loai
        counters[loai] += 1
        ngay: date = gd["ngay"]
        prefix = cfg["so_ct_thu"] if loai == "thu" else cfg["so_ct_chi"]
        so_ct = f"{prefix}{ngay.strftime('%y%m%d')}-{counters[loai]:03d}"
        so_tien = gd["tien_vao"] if loai == "thu" else gd["tien_ra"]
        tk_nh = tab_cfg.get("tk_hach_toan", "1121")
        tk_no, tk_co = (tk_nh, pl.tk_doi_ung) if loai == "thu" else (pl.tk_doi_ung, tk_nh)
        ghi_chu = f"Độ tin cậy {pl.do_tin_cay:.0%}" + ("" if pl.do_tin_cay >= 0.7 else " - CẦN KIỂM TRA")
        row = [
            utils.fmt_date(ngay), utils.fmt_date(ngay), so_ct, tab_cfg.get("so_tai_khoan", ""),
            pl.ma_doi_tuong, pl.ten_doi_tuong, pl.dien_giai, gd["noi_dung"],
            tk_no, tk_co, so_tien, _ten_cong_trinh(pl.cong_trinh_ma), _ten_khoan_muc(pl.khoan_muc_ma), ghi_chu,
        ]
        (rows_thu if loai == "thu" else rows_chi).append(row)
    return rows_thu, rows_chi


def xuat_file_misa(giao_dich: list[dict], phan_loai: list[GiaoDichNganHang], tab_cfg: dict, out_path: str | Path) -> Path:
    tpl = utils.misa_templates()
    rows_thu, rows_chi = build_rows(giao_dich, phan_loai, tab_cfg)
    wb = Workbook()
    wb.remove(wb.active)
    for key, rows in (("thu_tien_gui", rows_thu), ("chi_tien_gui", rows_chi)):
        ws = wb.create_sheet(tpl[key]["sheet"])
        cols = tpl[key]["cot"]
        ws.append(cols)
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = FILL_HEAD
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for r in rows:
            ws.append(r)
            if "CẦN KIỂM TRA" in str(r[-1]):
                for c in ws[ws.max_row]:
                    c.fill = FILL_LOW
        for i, w in enumerate([12, 12, 16, 16, 14, 30, 40, 50, 8, 8, 16, 24, 28, 22], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2, min_col=11, max_col=11):
            for c in row:
                c.number_format = "#,##0"
        ws.freeze_panes = "A2"
    out_path = Path(out_path)
    utils.ensure_dir(out_path.parent)
    wb.save(out_path)
    return out_path
