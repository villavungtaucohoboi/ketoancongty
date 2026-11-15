"""Đọc sao kê ngân hàng (từ Google Sheet, file Excel hoặc CSV tải từ ngân hàng) thành danh sách giao dịch chuẩn.

Bộ đọc tự tìm dòng tiêu đề (dòng có 'Nội dung giao dịch' / 'Ngày giao dịch' ...) nên chấp nhận được
nhiều mẫu sao kê khác nhau (ACB, Vietinbank, BIDV, MB...) miễn là có các cột ngày / nội dung / tiền ra / tiền vào.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from .. import utils

# Các tên cột có thể gặp -> tên chuẩn
COL_ALIASES = {
    "ngay": ["ngày giao dịch", "ngay giao dich", "ngày gd", "ngay gd", "transaction date", "ngày", "ngay hieu luc", "ngày hiệu lực", "ngày hạch toán", "value date", "trans date", "date"],
    "so_gd": ["số gd", "so gd", "số giao dịch", "so giao dich", "số bút toán", "mã giao dịch", "ref", "reference", "transaction no", "số ct", "so ct"],
    "noi_dung": ["nội dung giao dịch", "noi dung giao dich", "nội dung", "noi dung", "diễn giải", "dien giai", "description", "remark", "mô tả", "mo ta", "transaction description", "details"],
    "tien_ra": ["số tiền rút ra", "so tien rut ra", "tiền ra", "tien ra", "ghi nợ", "ghi no", "debit", "rút", "chi", "phát sinh nợ", "withdrawal", "số tiền ghi nợ"],
    "tien_vao": ["số tiền gửi vào", "so tien gui vao", "tiền vào", "tien vao", "ghi có", "ghi co", "credit", "gửi", "thu", "phát sinh có", "deposit", "số tiền ghi có"],
    "so_du": ["số dư", "so du", "balance", "số dư cuối"],
    "ma_cong_trinh_sheet": ["mã công trình", "ma cong trinh", "công trình", "cong trinh", "mã ct", "ma ct", "project"],
    "thang_nam": ["tháng - năm", "thang - nam", "tháng năm", "kỳ"],
}


def _norm(s: Any) -> str:
    return utils.bo_dau(str(s)).lower() if s is not None else ""


def _find_header_row(rows: list[list[Any]]) -> int:
    """Trả về chỉ số dòng tiêu đề: dòng có >= 3 cột khớp alias."""
    best, best_i = 0, -1
    for i, row in enumerate(rows[:60]):
        cells = [_norm(c) for c in row]
        hit = 0
        for aliases in COL_ALIASES.values():
            if any(any(_norm(a) == c or (_norm(a) in c and len(c) < 40) for c in cells if c) for a in aliases):
                hit += 1
        if hit > best:
            best, best_i = hit, i
    if best < 3:
        raise ValueError("Không tìm thấy dòng tiêu đề sao kê (cần có các cột Ngày / Nội dung / Tiền ra / Tiền vào).")
    return best_i


def _map_columns(header: list[Any]) -> dict[str, int]:
    cols: dict[str, int] = {}
    hn = [_norm(h) for h in header]
    for key, aliases in COL_ALIASES.items():
        for a in aliases:
            a = _norm(a)
            # ưu tiên khớp chính xác, rồi mới khớp chứa
            exact = [i for i, h in enumerate(hn) if h == a and i not in cols.values()]
            if exact:
                cols[key] = exact[0]
                break
        if key not in cols:
            for a in aliases:
                a = _norm(a)
                part = [i for i, h in enumerate(hn) if h and a in h and i not in cols.values()]
                if part:
                    cols[key] = part[0]
                    break
    for must in ("ngay", "noi_dung"):
        if must not in cols:
            raise ValueError(f"Thiếu cột bắt buộc '{must}' trong sao kê. Tiêu đề đọc được: {header}")
    if "tien_ra" not in cols and "tien_vao" not in cols:
        raise ValueError("Thiếu cột tiền ra / tiền vào trong sao kê.")
    return cols


def doc_sao_ke_tu_rows(rows: list[list[Any]]) -> list[dict]:
    """rows: toàn bộ ô của sheet (list of list). Trả về list giao dịch chuẩn."""
    h = _find_header_row(rows)
    cols = _map_columns(rows[h])
    out = []
    stt = 0
    for raw in rows[h + 1:]:
        if not raw or all(c in (None, "") for c in raw):
            continue
        get = lambda k: raw[cols[k]] if k in cols and cols[k] < len(raw) else None  # noqa: E731
        ngay = utils.parse_date(get("ngay"))
        nd = str(get("noi_dung") or "").strip()
        tien_ra = utils.parse_money(get("tien_ra"))
        tien_vao = utils.parse_money(get("tien_vao"))
        if ngay is None or (tien_ra == 0 and tien_vao == 0):
            continue  # dòng tổng, dòng trống, dòng tiêu đề phụ
        stt += 1
        out.append({
            "stt": stt,
            "ngay": ngay,
            "so_gd": str(get("so_gd") or "").strip(),
            "noi_dung": nd,
            "tien_ra": tien_ra,
            "tien_vao": tien_vao,
            "so_du": utils.parse_money(get("so_du")),
            "ma_cong_trinh_sheet": str(get("ma_cong_trinh_sheet") or "").strip(),
        })
    return out


def doc_sao_ke_tu_file(path: str | Path, sheet: str | int | None = None) -> list[dict]:
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        df = pd.read_csv(path, header=None, dtype=str, encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(path, header=None, sheet_name=sheet if sheet is not None else 0)
    rows = df.where(pd.notnull(df), None).values.tolist()
    return doc_sao_ke_tu_rows(rows)
