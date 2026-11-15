"""Tính lương hàng tháng từ: bảng nhân sự (lương cơ bản, phụ cấp, số người phụ thuộc)
+ bảng chấm công (ngày công, tăng ca) + bảng đánh giá (xếp loại) theo quy tắc trong config/luong_rules.yaml.

Đây là ENGINE MẪU: công thức được viết theo cách phổ biến ở doanh nghiệp Việt Nam. Khi nhận được
bảng lương mẫu thật của công ty, chỉ cần sửa luong_rules.yaml và (nếu cần) hàm tinh_mot_nguoi().
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .. import utils

COL_NS = {  # tên cột chuẩn -> các tên có thể gặp trong file nhân sự
    "ma_nv": ["mã nv", "ma nv", "mã nhân viên", "mã"],
    "ho_ten": ["họ tên", "ho ten", "họ và tên", "tên"],
    "bo_phan": ["bộ phận", "bo phan", "phòng ban"],
    "luong_cb": ["lương cơ bản", "luong co ban", "lương cb", "lương thỏa thuận", "mức lương"],
    "luong_bh": ["lương đóng bh", "lương đóng bảo hiểm", "luong dong bh", "lương bhxh"],
    "pc_an_trua": ["phụ cấp ăn trưa", "ăn trưa", "pc ăn"],
    "pc_dien_thoai": ["phụ cấp điện thoại", "điện thoại"],
    "pc_xang_xe": ["phụ cấp xăng xe", "xăng xe", "đi lại"],
    "pc_trach_nhiem": ["phụ cấp trách nhiệm", "trách nhiệm", "pc chức vụ"],
    "so_npt": ["số người phụ thuộc", "người phụ thuộc", "npt"],
    "tam_ung": ["tạm ứng", "tam ung"],
}
COL_CC = {
    "ma_nv": ["mã nv", "ma nv", "mã nhân viên", "mã"],
    "ngay_cong": ["ngày công", "ngay cong", "công thực tế", "số công", "tổng công"],
    "ngay_phep": ["phép", "nghỉ phép", "phép có lương"],
    "ot_thuong": ["tăng ca ngày thường", "ot ngày thường", "giờ tăng ca", "tăng ca"],
    "ot_nghi": ["tăng ca ngày nghỉ", "ot ngày nghỉ", "ot cuối tuần"],
    "ot_le": ["tăng ca ngày lễ", "ot lễ"],
}
COL_DG = {
    "ma_nv": ["mã nv", "ma nv", "mã nhân viên", "mã"],
    "xep_loai": ["xếp loại", "xep loai", "đánh giá", "kết quả", "loại"],
    "thuong_khac": ["thưởng khác", "thưởng", "thưởng nóng"],
    "phat": ["phạt", "trừ", "khấu trừ khác"],
}


def _norm(s) -> str:
    return utils.bo_dau(str(s)).lower() if s is not None else ""


def _chuan_hoa(df: pd.DataFrame, mapping: dict[str, list[str]]) -> pd.DataFrame:
    cols = {}
    hn = {c: _norm(c) for c in df.columns}
    for key, al in mapping.items():
        for a in al:
            a = _norm(a)
            hit = [c for c, h in hn.items() if h == a and c not in cols.values()]
            if hit:
                cols[key] = hit[0]
                break
        if key not in cols:
            for a in al:
                a = _norm(a)
                hit = [c for c, h in hn.items() if a in h and c not in cols.values()]
                if hit:
                    cols[key] = hit[0]
                    break
    out = pd.DataFrame({k: df[v] for k, v in cols.items()})
    if "ma_nv" not in out:
        raise ValueError(f"Thiếu cột 'Mã NV' trong file. Các cột đọc được: {list(df.columns)}")
    out["ma_nv"] = out["ma_nv"].astype(str).str.strip()
    return out


def _doc(path, mapping) -> pd.DataFrame:
    df = pd.read_excel(path) if str(path).lower().endswith((".xlsx", ".xls")) else pd.read_csv(path, encoding="utf-8-sig")
    return _chuan_hoa(df, mapping)


def thue_tncn(thu_nhap_tinh_thue: float, bieu: list) -> float:
    if thu_nhap_tinh_thue <= 0:
        return 0.0
    thue, truoc = 0.0, 0.0
    for tran, suat in bieu:
        if tran is None or thu_nhap_tinh_thue <= tran:
            thue += (thu_nhap_tinh_thue - truoc) * suat
            break
        thue += (tran - truoc) * suat
        truoc = tran
    return round(thue)


def tinh_mot_nguoi(ns: dict, cc: dict, dg: dict, r: dict) -> dict:
    f = lambda d, k: utils.parse_money(d.get(k)) if d.get(k) not in (None, "") and not pd.isna(d.get(k)) else 0.0  # noqa: E731
    ngay_chuan = float(r.get("ngay_cong_chuan", 26))
    luong_cb = f(ns, "luong_cb")
    luong_bh = f(ns, "luong_bh") or luong_cb
    ngay_cong = f(cc, "ngay_cong") + f(cc, "ngay_phep")
    luong_ngay = luong_cb / ngay_chuan if ngay_chuan else 0
    luong_gio = luong_ngay / 8
    luong_theo_cong = round(luong_ngay * ngay_cong)
    pc_an = f(ns, "pc_an_trua")
    pc_dt = f(ns, "pc_dien_thoai")
    pc_xx = f(ns, "pc_xang_xe")
    pc_tn = f(ns, "pc_trach_nhiem")
    hs = r.get("he_so_tang_ca", {})
    tien_ot = round(luong_gio * (f(cc, "ot_thuong") * hs.get("ngay_thuong", 1.5)
                                 + f(cc, "ot_nghi") * hs.get("ngay_nghi", 2.0)
                                 + f(cc, "ot_le") * hs.get("ngay_le", 3.0)))
    # Phần tăng ca vượt mức giờ thường (ví dụ 0.5 của 1.5) được miễn thuế TNCN
    ot_mien_thue = round(luong_gio * (f(cc, "ot_thuong") * (hs.get("ngay_thuong", 1.5) - 1)
                                      + f(cc, "ot_nghi") * (hs.get("ngay_nghi", 2.0) - 1)
                                      + f(cc, "ot_le") * (hs.get("ngay_le", 3.0) - 1)))
    xep_loai = str(dg.get("xep_loai") or "").strip().upper()
    ty_le_thuong = float(r.get("thuong_theo_xep_loai", {}).get(xep_loai, 0.0))
    thuong_kpi = round(luong_cb * ty_le_thuong)
    thuong_khac = f(dg, "thuong_khac")
    phat = f(dg, "phat")
    tong_thu_nhap = luong_theo_cong + pc_an + pc_dt + pc_xx + pc_tn + tien_ot + thuong_kpi + thuong_khac - phat

    bh = r.get("bao_hiem_nld", {})
    can_cu_bhxh = min(luong_bh, float(r.get("tran_luong_bhxh", 1e18)))
    can_cu_bhtn = min(luong_bh, float(r.get("tran_luong_bhtn", 1e18)))
    bhxh = round(can_cu_bhxh * bh.get("bhxh", 0.08))
    bhyt = round(can_cu_bhxh * bh.get("bhyt", 0.015))
    bhtn = round(can_cu_bhtn * bh.get("bhtn", 0.01))
    tong_bh = bhxh + bhyt + bhtn
    bhdn = r.get("bao_hiem_dn", {})
    bh_dn = round(can_cu_bhxh * (bhdn.get("bhxh", 0) + bhdn.get("bhyt", 0) + bhdn.get("bhtnld", 0)) + can_cu_bhtn * bhdn.get("bhtn", 0))

    mien_an = min(pc_an, float(r.get("muc_mien_an_trua", 730000)))
    thu_nhap_chiu_thue = tong_thu_nhap - mien_an - pc_dt - pc_xx - ot_mien_thue
    gt = r.get("giam_tru", {})
    giam_tru = float(gt.get("ban_than", 0)) + float(gt.get("nguoi_phu_thuoc", 0)) * f(ns, "so_npt")
    tn_tinh_thue = max(0.0, thu_nhap_chiu_thue - tong_bh - giam_tru)
    thue = thue_tncn(tn_tinh_thue, r.get("bieu_thue", []))
    tam_ung = f(ns, "tam_ung")
    thuc_linh = round(tong_thu_nhap - tong_bh - thue - tam_ung)
    return {
        "Mã NV": ns.get("ma_nv"), "Họ tên": ns.get("ho_ten", ""), "Bộ phận": ns.get("bo_phan", ""),
        "Lương cơ bản": luong_cb, "Ngày công": ngay_cong, "Lương theo công": luong_theo_cong,
        "PC ăn trưa": pc_an, "PC điện thoại": pc_dt, "PC xăng xe": pc_xx, "PC trách nhiệm": pc_tn,
        "Tiền tăng ca": tien_ot, "Xếp loại": xep_loai, "Thưởng KPI": thuong_kpi, "Thưởng khác": thuong_khac, "Phạt/trừ": phat,
        "Tổng thu nhập": tong_thu_nhap, "BHXH (8%)": bhxh, "BHYT (1.5%)": bhyt, "BHTN (1%)": bhtn, "Tổng BH NLĐ": tong_bh,
        "Giảm trừ gia cảnh": giam_tru, "TN tính thuế": tn_tinh_thue, "Thuế TNCN": thue, "Tạm ứng": tam_ung,
        "THỰC LĨNH": thuc_linh, "BH DN đóng": bh_dn,
    }


def tinh_luong(file_nhan_su, file_cham_cong, file_danh_gia=None, rules: dict | None = None) -> pd.DataFrame:
    r = rules or utils.load_yaml("luong_rules.yaml")
    ns = _doc(file_nhan_su, COL_NS)
    cc = _doc(file_cham_cong, COL_CC)
    dg = _doc(file_danh_gia, COL_DG) if file_danh_gia else pd.DataFrame(columns=["ma_nv"])
    cc_i = cc.set_index("ma_nv")
    dg_i = dg.set_index("ma_nv") if not dg.empty else dg
    rows = []
    for _, n in ns.iterrows():
        ma = n["ma_nv"]
        c = cc_i.loc[ma].to_dict() if ma in cc_i.index else {}
        d = dg_i.loc[ma].to_dict() if (len(dg_i) and ma in dg_i.index) else {}
        if isinstance(c.get("ngay_cong"), pd.Series):  # trùng mã
            c = {k: (v.iloc[0] if isinstance(v, pd.Series) else v) for k, v in c.items()}
        rows.append(tinh_mot_nguoi(n.to_dict(), c, d, r))
    return pd.DataFrame(rows)


def xuat_bang_luong(df: pd.DataFrame, out_path, thang: str = "") -> Path:
    """Bảng thanh toán tiền lương theo mẫu kế toán Việt Nam: tiêu đề công ty, nhóm cột, dòng tổng (công thức), chữ ký."""
    from openpyxl.styles import Alignment, Border, Side
    from openpyxl.utils import get_column_letter as L

    out_path = Path(out_path)
    utils.ensure_dir(out_path.parent)
    cty = utils.settings().get("cong_ty", {})
    thang_txt = thang or ""
    try:
        m, y = thang.split("/")
        thang_txt = f"Tháng {int(m):02d} năm {y}"
    except Exception:  # noqa: BLE001
        pass

    FONT = "Times New Roman"
    f_norm = Font(name=FONT, size=11)
    f_bold = Font(name=FONT, size=11, bold=True)
    f_title = Font(name=FONT, size=16, bold=True)
    f_head = Font(name=FONT, size=10, bold=True)
    f_small = Font(name=FONT, size=10, italic=True)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_head = PatternFill("solid", fgColor="DDEBF7")
    fill_group = PatternFill("solid", fgColor="BDD7EE")
    fill_total = PatternFill("solid", fgColor="FFF2CC")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    # ---- cấu trúc cột: (nhóm, tiêu đề, cột nguồn trong df | None nếu là công thức, rộng)
    cols = [
        ("", "STT", None, 5), ("", "Mã NV", "Mã NV", 9), ("", "Họ và tên", "Họ tên", 22), ("", "Bộ phận", "Bộ phận", 12),
        ("", "Lương cơ bản", "Lương cơ bản", 13), ("", "Ngày công", "Ngày công", 8), ("", "Lương theo công", "Lương theo công", 14),
        ("Phụ cấp", "Ăn trưa", "PC ăn trưa", 11), ("Phụ cấp", "Điện thoại", "PC điện thoại", 11), ("Phụ cấp", "Xăng xe", "PC xăng xe", 11), ("Phụ cấp", "Trách nhiệm", "PC trách nhiệm", 12),
        ("", "Tiền tăng ca", "Tiền tăng ca", 12), ("Thưởng", "Xếp loại", "Xếp loại", 7), ("Thưởng", "KPI", "Thưởng KPI", 12), ("Thưởng", "Khác", "Thưởng khác", 11),
        ("", "Phạt / trừ", "Phạt/trừ", 10), ("", "TỔNG THU NHẬP", "=TONG_TN", 15),
        ("Các khoản khấu trừ", "BHXH (8%)", "BHXH (8%)", 11), ("Các khoản khấu trừ", "BHYT (1,5%)", "BHYT (1.5%)", 11), ("Các khoản khấu trừ", "BHTN (1%)", "BHTN (1%)", 11),
        ("Các khoản khấu trừ", "Thuế TNCN", "Thuế TNCN", 12), ("Các khoản khấu trừ", "Tạm ứng", "Tạm ứng", 11), ("Các khoản khấu trừ", "Cộng khấu trừ", "=TONG_KT", 13),
        ("", "THỰC LĨNH", "=THUC_LINH", 15), ("", "Ký nhận", None, 12),
    ]
    n = len(cols)
    wb = Workbook()
    ws = wb.active
    ws.title = "Bang luong"
    ws.sheet_view.showGridLines = False

    ws["A1"] = cty.get("ten", "CÔNG TY")
    ws["A1"].font = f_bold
    ws["A2"] = "Địa chỉ: " + cty.get("dia_chi", "..........")
    ws["A2"].font = f_norm
    ws.cell(row=1, column=n - 3, value="Mẫu số: 02-LĐTL").font = f_small
    ws.cell(row=2, column=n - 3, value="(Ban hành theo Thông tư 133/2016/TT-BTC)").font = f_small
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=n)
    ws.cell(row=4, column=1, value="BẢNG THANH TOÁN TIỀN LƯƠNG").font = f_title
    ws.cell(row=4, column=1).alignment = center
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=n)
    ws.cell(row=5, column=1, value=thang_txt).font = Font(name=FONT, size=12, italic=True)
    ws.cell(row=5, column=1).alignment = center
    ws.cell(row=6, column=n - 1, value="Đơn vị tính: VNĐ").font = f_small
    ws.cell(row=6, column=n - 1).alignment = right

    # ---- 2 dòng tiêu đề nhóm
    H1, H2 = 7, 8
    c = 1
    while c <= n:
        g = cols[c - 1][0]
        if g:
            e = c
            while e < n and cols[e][0] == g:
                e += 1
            ws.merge_cells(start_row=H1, start_column=c, end_row=H1, end_column=e)
            ws.cell(row=H1, column=c, value=g)
            for cc in range(c, e + 1):
                ws.cell(row=H2, column=cc, value=cols[cc - 1][1])
            c = e + 1
        else:
            ws.merge_cells(start_row=H1, start_column=c, end_row=H2, end_column=c)
            ws.cell(row=H1, column=c, value=cols[c - 1][1])
            c += 1
    for r in (H1, H2):
        for cc in range(1, n + 1):
            cell = ws.cell(row=r, column=cc)
            cell.font = f_head
            cell.alignment = center
            cell.border = border
            cell.fill = fill_group if (r == H1 and cols[cc - 1][0]) else fill_head
    ws.row_dimensions[H1].height = 22
    ws.row_dimensions[H2].height = 30
    # dòng số thứ tự cột (A, B, C...) theo kiểu biểu mẫu
    R0 = H2 + 1
    for cc in range(1, n + 1):
        cell = ws.cell(row=R0, column=cc, value=cc)
        cell.font = f_small
        cell.alignment = center
        cell.border = border

    idx = {t: j + 1 for j, (_, t, _, _) in enumerate(cols)}
    col_of = {src: j + 1 for j, (_, _, src, _) in enumerate(cols) if src and not str(src).startswith("=")}
    first = R0 + 1
    for i, rec in enumerate(df.to_dict("records")):
        r = first + i
        for j, (_, title, src, _) in enumerate(cols, start=1):
            cell = ws.cell(row=r, column=j)
            if title == "STT":
                cell.value = i + 1
                cell.alignment = center
            elif title == "Ký nhận":
                cell.value = None
            elif src == "=TONG_TN":
                cell.value = (f"=SUM({L(col_of['Lương theo công'])}{r}:{L(col_of['Tiền tăng ca'])}{r})"
                              f"+{L(col_of['Thưởng KPI'])}{r}+{L(col_of['Thưởng khác'])}{r}-{L(col_of['Phạt/trừ'])}{r}")
            elif src == "=TONG_KT":
                cell.value = f"=SUM({L(col_of['BHXH (8%)'])}{r}:{L(col_of['Tạm ứng'])}{r})"
            elif src == "=THUC_LINH":
                cell.value = f"={L(idx['TỔNG THU NHẬP'])}{r}-{L(idx['Cộng khấu trừ'])}{r}"
            else:
                v = rec.get(src)
                cell.value = v if not (isinstance(v, float) and pd.isna(v)) else None
            cell.font = f_bold if title in ("TỔNG THU NHẬP", "THỰC LĨNH") else f_norm
            cell.border = border
            if title in ("Mã NV", "Họ và tên", "Bộ phận"):
                cell.alignment = left
            elif title in ("Xếp loại", "Ngày công", "STT"):
                cell.alignment = center
            else:
                cell.alignment = right
                cell.number_format = "#,##0"
    last = first + len(df) - 1
    # ---- dòng tổng (công thức SUM)
    T = last + 1
    ws.merge_cells(start_row=T, start_column=1, end_row=T, end_column=4)
    ws.cell(row=T, column=1, value="TỔNG CỘNG").font = f_bold
    ws.cell(row=T, column=1).alignment = center
    for j, (_, title, src, _) in enumerate(cols, start=1):
        cell = ws.cell(row=T, column=j)
        if j >= 5 and title not in ("Xếp loại", "Ký nhận"):
            cell.value = f"=SUM({L(j)}{first}:{L(j)}{last})"
            cell.number_format = "#,##0" if title != "Ngày công" else "0.##"
            cell.alignment = right
        cell.font = f_bold
        cell.border = border
        cell.fill = fill_total
    # ---- chữ ký
    S = T + 2
    ws.cell(row=S, column=n - 5, value=f"Ngày ...... tháng ...... năm ......").font = f_small
    ws.merge_cells(start_row=S, start_column=n - 5, end_row=S, end_column=n)
    ws.cell(row=S, column=n - 5).alignment = center
    for col, ten in ((2, "Người lập biểu"), (int(n / 2) - 2, "Kế toán trưởng"), (n - 5, "Giám đốc")):
        ws.merge_cells(start_row=S + 1, start_column=col, end_row=S + 1, end_column=col + 5)
        ws.merge_cells(start_row=S + 2, start_column=col, end_row=S + 2, end_column=col + 5)
        ws.cell(row=S + 1, column=col, value=ten).font = f_bold
        ws.cell(row=S + 1, column=col).alignment = center
        ws.cell(row=S + 2, column=col, value="(Ký, họ tên)").font = f_small
        ws.cell(row=S + 2, column=col).alignment = center
    for j, (_, _, _, w) in enumerate(cols, start=1):
        ws.column_dimensions[L(j)].width = w
    ws.freeze_panes = ws.cell(row=first, column=4)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{H1}:{H2}"

    # ---- sheet 2: chi tiết thuế TNCN
    ws2 = wb.create_sheet("Chi tiet thue TNCN")
    cols2 = ["STT", "Mã NV", "Họ tên", "Tổng thu nhập", "Tổng BH NLĐ", "Giảm trừ gia cảnh", "TN tính thuế", "Thuế TNCN", "BH DN đóng"]
    ws2.append(cols2)
    for cell in ws2[1]:
        cell.font = f_head
        cell.fill = fill_head
        cell.border = border
        cell.alignment = center
    for i, rec in enumerate(df.to_dict("records"), start=1):
        ws2.append([i] + [rec.get(c) for c in cols2[1:]])
        for cell in ws2[ws2.max_row]:
            cell.font = f_norm
            cell.border = border
            if cell.column >= 4:
                cell.number_format = "#,##0"
    for j, w in enumerate([5, 9, 24, 15, 14, 16, 15, 13, 13], start=1):
        ws2.column_dimensions[L(j)].width = w
    ws2.cell(row=ws2.max_row + 2, column=1, value="Giảm trừ, tỷ lệ BH và biểu thuế lấy từ config/luong_rules.yaml; kế toán kiểm tra lại mức áp dụng cho kỳ hiện hành.").font = f_small
    wb.save(out_path)
    return out_path
