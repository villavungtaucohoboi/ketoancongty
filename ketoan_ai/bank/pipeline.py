"""Luồng xử lý sao kê: đọc -> phân loại -> xuất file nhập MISA + bảng xem trước."""
from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

from .. import ai, utils
from . import misa_export, parser, sheets


def chay(tab: str, nguon: str | Path | None = None, tu_ngay: date | None = None, den_ngay: date | None = None,
         out_dir: str | Path | None = None) -> tuple[pd.DataFrame, Path]:
    """tab: tên tab trong settings.yaml (ACB_CTY...). nguon: đường dẫn file xlsx/csv; None = đọc Google Sheet."""
    cfg = utils.settings()["sao_ke"]
    tab_cfg = cfg["tabs"].get(tab)
    if tab_cfg is None:
        raise KeyError(f"Tab '{tab}' chưa khai báo trong config/settings.yaml > sao_ke.tabs")
    if nguon:
        gds = parser.doc_sao_ke_tu_file(nguon)
    else:
        sheet_id = utils.env("KT_SHEET_ID")
        if not sheet_id:
            raise RuntimeError("Chưa cấu hình KT_SHEET_ID trong .env")
        gds = parser.doc_sao_ke_tu_rows(sheets.doc_tab(sheet_id, tab))
    if tu_ngay:
        gds = [g for g in gds if g["ngay"] >= tu_ngay]
    if den_ngay:
        gds = [g for g in gds if g["ngay"] <= den_ngay]
    for i, g in enumerate(gds, start=1):
        g["stt"] = i
    if not gds:
        raise ValueError("Không có giao dịch nào trong khoảng ngày đã chọn.")
    pl = ai.phan_loai_giao_dich(gds)
    out_dir = Path(out_dir) if out_dir else utils.DATA_DIR / "misa_import"
    ten = f"MISA_NganHang_{tab}_{gds[0]['ngay']:%Y%m%d}_{gds[-1]['ngay']:%Y%m%d}.xlsx"
    out = misa_export.xuat_file_misa(gds, pl, tab_cfg, out_dir / ten)
    df = pd.DataFrame([{
        "Ngày": utils.fmt_date(g["ngay"]), "Loại": p.loai.upper(), "Số tiền": g["tien_vao"] or g["tien_ra"],
        "Nội dung": g["noi_dung"], "TK đối ứng": p.tk_doi_ung, "Đối tượng": p.ma_doi_tuong or p.ten_doi_tuong,
        "Công trình": p.cong_trinh_ma, "Khoản mục": p.khoan_muc_ma, "Diễn giải": p.dien_giai, "Tin cậy": p.do_tin_cay,
    } for g, p in zip(gds, pl)])
    return df, out
