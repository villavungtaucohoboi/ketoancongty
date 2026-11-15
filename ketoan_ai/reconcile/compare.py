"""Đối chiếu bảng kê hóa đơn mua vào: phần mềm MISA  <->  bảng kê tải từ Cổng HĐĐT (hoadondientu.gdt.gov.vn).

Khóa đối chiếu: MST người bán + Số hóa đơn (+ Ký hiệu nếu cả 2 file đều có).
Kết quả: 4 nhóm - Khớp / Lệch tiền / Chỉ có trên Thuế (chưa nhập phần mềm) / Chỉ có trên phần mềm (thuế chưa có).
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .. import utils

ALIASES = {
    "so_hd": ["số hóa đơn", "so hoa don", "số hđ", "so hd", "số", "invoice no"],
    "ky_hieu": ["ký hiệu", "ky hieu", "ký hiệu hóa đơn", "ký hiệu hđ", "kí hiệu"],
    "mau_so": ["mẫu số", "mau so", "ký hiệu mẫu số"],
    "ngay": ["ngày hóa đơn", "ngay hoa don", "ngày lập", "ngay lap", "ngày hđ", "ngày"],
    "mst": ["mst người bán", "mã số thuế người bán", "mst nguoi ban", "mst", "mã số thuế", "ma so thue", "mst nb"],
    "ten": ["tên người bán", "ten nguoi ban", "người bán", "nhà cung cấp", "tên nhà cung cấp", "đơn vị bán", "tên đơn vị"],
    "tien_hang": ["tổng tiền chưa thuế", "tien chua thue", "doanh số mua chưa có thuế", "giá trị hhdv mua vào chưa thuế", "tiền hàng", "thành tiền", "giá trị chưa thuế", "tổng tiền hàng"],
    "tien_thue": ["tổng tiền thuế", "tiền thuế gtgt", "thuế gtgt", "tien thue", "tổng tiền thuế gtgt"],
    "tong": ["tổng tiền thanh toán", "tong tien thanh toan", "tổng cộng", "tổng thanh toán", "tổng tiền"],
    "trang_thai": ["trạng thái hóa đơn", "trạng thái", "tt hóa đơn"],
}


def _norm(s) -> str:
    return utils.bo_dau(str(s)).lower() if s is not None else ""


def _doc_bang_ke(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        raw = pd.read_csv(path, header=None, dtype=str, encoding="utf-8-sig", sep=None, engine="python")
    else:
        raw = pd.read_excel(path, header=None, dtype=str)
    rows = raw.where(pd.notnull(raw), None).values.tolist()
    # tìm dòng tiêu đề
    best, hi = 0, 0
    for i, r in enumerate(rows[:40]):
        cells = [_norm(c) for c in r]
        hit = sum(1 for al in ALIASES.values() if any(any(_norm(a) == c or _norm(a) in c for c in cells if c) for a in al))
        if hit > best:
            best, hi = hit, i
    if best < 3:
        raise ValueError(f"Không nhận ra tiêu đề bảng kê trong {path.name}")
    header = [_norm(c) for c in rows[hi]]
    colmap: dict[str, int] = {}
    for key, al in ALIASES.items():
        for a in al:
            a = _norm(a)
            idx = [i for i, h in enumerate(header) if h == a and i not in colmap.values()]
            if idx:
                colmap[key] = idx[0]
                break
        if key not in colmap:
            for a in al:
                a = _norm(a)
                idx = [i for i, h in enumerate(header) if h and a in h and i not in colmap.values()]
                if idx:
                    colmap[key] = idx[0]
                    break
    recs = []
    for r in rows[hi + 1:]:
        if not r or all(c in (None, "") for c in r):
            continue
        g = lambda k: (r[colmap[k]] if k in colmap and colmap[k] < len(r) else None)  # noqa: E731
        so = str(g("so_hd") or "").strip()
        if not so or not re.search(r"\d", so):
            continue
        recs.append({
            "so_hd": so.lstrip("0") or "0",
            "ky_hieu": str(g("ky_hieu") or "").strip().upper(),
            "mau_so": str(g("mau_so") or "").strip(),
            "ngay": utils.parse_date(g("ngay")),
            "mst": re.sub(r"\D", "", str(g("mst") or "")),
            "ten": str(g("ten") or "").strip(),
            "tien_hang": utils.parse_money(g("tien_hang")),
            "tien_thue": utils.parse_money(g("tien_thue")),
            "tong": utils.parse_money(g("tong")),
            "trang_thai": str(g("trang_thai") or "").strip(),
        })
    df = pd.DataFrame(recs)
    if df.empty:
        return df
    # Bảng kê MISA có thể nhiều dòng / 1 hóa đơn (theo mặt hàng) -> gộp
    df = (df.groupby(["mst", "so_hd"], as_index=False)
            .agg(ky_hieu=("ky_hieu", "first"), mau_so=("mau_so", "first"), ngay=("ngay", "first"), ten=("ten", "first"),
                 tien_hang=("tien_hang", "sum"), tien_thue=("tien_thue", "sum"), tong=("tong", "sum"), trang_thai=("trang_thai", "first")))
    df["tong"] = df.apply(lambda x: x["tong"] if x["tong"] else x["tien_hang"] + x["tien_thue"], axis=1)
    return df


def doi_chieu(file_misa: str | Path, file_thue: str | Path) -> dict[str, pd.DataFrame]:
    tol = float(utils.settings().get("doi_chieu", {}).get("dung_sai_tien", 1000))
    m = _doc_bang_ke(file_misa)
    t = _doc_bang_ke(file_thue)
    if m.empty or t.empty:
        raise ValueError("Một trong hai bảng kê không có dữ liệu hóa đơn.")
    key = ["mst", "so_hd"]
    j = m.merge(t, on=key, how="outer", suffixes=("_misa", "_thue"), indicator=True)

    def _fmt(df: pd.DataFrame, side: str) -> pd.DataFrame:
        return pd.DataFrame({
            "MST người bán": df["mst"], "Tên người bán": df[f"ten_{side}"], "Ký hiệu": df[f"ky_hieu_{side}"],
            "Số HĐ": df["so_hd"], "Ngày HĐ": df[f"ngay_{side}"].map(utils.fmt_date),
            "Tiền hàng": df[f"tien_hang_{side}"], "Tiền thuế": df[f"tien_thue_{side}"], "Tổng": df[f"tong_{side}"],
            "Trạng thái (thuế)": df.get("trang_thai_thue", ""),
        })

    both = j[j["_merge"] == "both"].copy()
    both["lech_hang"] = both["tien_hang_misa"] - both["tien_hang_thue"]
    both["lech_thue"] = both["tien_thue_misa"] - both["tien_thue_thue"]
    both["lech_tong"] = both["tong_misa"] - both["tong_thue"]
    khop_mask = (both["lech_hang"].abs() <= tol) & (both["lech_thue"].abs() <= tol) & (both["lech_tong"].abs() <= tol)
    khop = both[khop_mask]
    lech = both[~khop_mask]
    lech_df = pd.DataFrame({
        "MST người bán": lech["mst"], "Tên người bán": lech["ten_thue"], "Ký hiệu": lech["ky_hieu_thue"], "Số HĐ": lech["so_hd"],
        "Tiền hàng MISA": lech["tien_hang_misa"], "Tiền hàng Thuế": lech["tien_hang_thue"], "Lệch tiền hàng": lech["lech_hang"],
        "Thuế MISA": lech["tien_thue_misa"], "Thuế Thuế": lech["tien_thue_thue"], "Lệch thuế": lech["lech_thue"],
        "Tổng MISA": lech["tong_misa"], "Tổng Thuế": lech["tong_thue"], "Lệch tổng": lech["lech_tong"],
    })
    chi_thue = _fmt(j[j["_merge"] == "right_only"], "thue")
    chi_misa = _fmt(j[j["_merge"] == "left_only"], "misa").drop(columns=["Trạng thái (thuế)"])
    tong_hop = pd.DataFrame([
        {"Chỉ tiêu": "Số HĐ trên MISA", "Số lượng": len(m), "Tiền hàng": m["tien_hang"].sum(), "Tiền thuế": m["tien_thue"].sum()},
        {"Chỉ tiêu": "Số HĐ trên Thuế", "Số lượng": len(t), "Tiền hàng": t["tien_hang"].sum(), "Tiền thuế": t["tien_thue"].sum()},
        {"Chỉ tiêu": "Khớp hoàn toàn", "Số lượng": len(khop), "Tiền hàng": khop["tien_hang_thue"].sum(), "Tiền thuế": khop["tien_thue_thue"].sum()},
        {"Chỉ tiêu": "Có cả 2 nhưng lệch tiền", "Số lượng": len(lech), "Tiền hàng": lech["tien_hang_thue"].sum(), "Tiền thuế": lech["tien_thue_thue"].sum()},
        {"Chỉ tiêu": "Chỉ có trên Thuế (chưa nhập MISA)", "Số lượng": len(chi_thue), "Tiền hàng": chi_thue["Tiền hàng"].sum(), "Tiền thuế": chi_thue["Tiền thuế"].sum()},
        {"Chỉ tiêu": "Chỉ có trên MISA (thuế chưa có / nhập sai)", "Số lượng": len(chi_misa), "Tiền hàng": chi_misa["Tiền hàng"].sum(), "Tiền thuế": chi_misa["Tiền thuế"].sum()},
    ])
    return {"Tong hop": tong_hop, "Khop": _fmt(khop, "thue"), "Lech tien": lech_df,
            "Chi co tren Thue": chi_thue, "Chi co tren MISA": chi_misa}


def xuat_bao_cao(ket_qua: dict[str, pd.DataFrame], out_path: str | Path) -> Path:
    out_path = Path(out_path)
    utils.ensure_dir(out_path.parent)
    wb = Workbook()
    wb.remove(wb.active)
    fills = {"Lech tien": "FCE4D6", "Chi co tren Thue": "FFF2CC", "Chi co tren MISA": "E2EFDA"}
    for name, df in ket_qua.items():
        ws = wb.create_sheet(name)
        ws.append(list(df.columns))
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor=fills.get(name, "D9E1F2"))
        for row in df.itertuples(index=False):
            ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
        for i, col in enumerate(df.columns, start=1):
            ws.column_dimensions[get_column_letter(i)].width = 18 if "Tên" not in col else 45
            if any(k in col for k in ("Tiền", "Thuế", "Tổng", "Lệch")):
                for cell in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    cell[0].number_format = "#,##0"
        ws.freeze_panes = "A2"
    wb.save(out_path)
    return out_path
