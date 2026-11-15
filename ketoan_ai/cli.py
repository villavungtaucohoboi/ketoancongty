"""Dòng lệnh: python -m ketoan_ai.cli <lenh> ...

  sao-ke   --tab ACB_CTY [--file sao_ke.xlsx] [--tu 01/01/2026 --den 12/01/2026]
  hoa-don  email [--ngay 7] [--khong-link]
  hoa-don  thu-muc <thu_muc> [--ghi-chu "Tên công trình"]
  hoa-don  xuat-misa
  doi-chieu --misa bang_ke_misa.xlsx --thue bang_ke_thue.xlsx
  luong    --nhan-su ns.xlsx --cham-cong cc.xlsx [--danh-gia dg.xlsx] [--thang 10/2026]
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime

from . import utils


def main(argv=None):
    p = argparse.ArgumentParser(prog="ketoan-ai", description="Trợ lý kế toán AI")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("sao-ke", help="Sao kê ngân hàng -> file nhập MISA")
    s.add_argument("--tab", required=True)
    s.add_argument("--file", help="File xlsx/csv sao kê (bỏ trống = đọc Google Sheet)")
    s.add_argument("--tu")
    s.add_argument("--den")

    h = sub.add_parser("hoa-don", help="Hóa đơn mua vào")
    hs = h.add_subparsers(dest="sub", required=True)
    he = hs.add_parser("email")
    he.add_argument("--ngay", type=int)
    he.add_argument("--khong-link", action="store_true")
    ht = hs.add_parser("thu-muc")
    ht.add_argument("thu_muc")
    ht.add_argument("--ghi-chu", default="")
    hs.add_parser("xuat-misa")
    hs.add_parser("danh-sach")

    d = sub.add_parser("doi-chieu", help="Đối chiếu bảng kê MISA với bảng kê thuế")
    d.add_argument("--misa", help="File bảng kê xuất từ MISA; bỏ trống = đọc thẳng CSDL MISA")
    d.add_argument("--thue", required=True)
    d.add_argument("--tu", help="Từ ngày dd/mm/yyyy (khi đọc thẳng CSDL MISA)")
    d.add_argument("--den", help="Đến ngày dd/mm/yyyy")

    m = sub.add_parser("misa", help="Kết nối CSDL MISA SME.NET (chỉ đọc)")
    m.add_argument("viec", choices=["kham-pha"])

    l = sub.add_parser("luong", help="Tính lương")
    l.add_argument("--nhan-su", required=True)
    l.add_argument("--cham-cong", required=True)
    l.add_argument("--danh-gia")
    l.add_argument("--thang", default=f"{datetime.now():%m/%Y}")

    a = p.parse_args(argv)
    if a.cmd == "sao-ke":
        from .bank import pipeline
        df, out = pipeline.chay(a.tab, a.file, utils.parse_date(a.tu), utils.parse_date(a.den))
        print(df.to_string(index=False, max_colwidth=40))
        print(f"\n=> File nhập MISA: {out}")
    elif a.cmd == "hoa-don":
        from .invoices import pipeline, store
        if a.sub == "email":
            tk = pipeline.chay_tu_email(a.ngay, dung_link=not a.khong_link)
        elif a.sub == "thu-muc":
            tk = pipeline.chay_tu_thu_muc(a.thu_muc, a.ghi_chu)
        elif a.sub == "xuat-misa":
            out, n = pipeline.xuat_misa()
            print(f"Đã xuất {n} hóa đơn -> {out}" if out else "Không có hóa đơn chờ xuất.")
            return
        else:
            for it in store.danh_sach():
                hd, pl = it["hd"], it["pl"]
                print(f"[{it['trang_thai']}] {hd['ngay_hoa_don']} {hd['ky_hieu']}/{hd['so_hoa_don']} {hd['ten_nguoi_ban'][:40]:40} "
                      f"{utils.fmt_money(hd['tong_thanh_toan']):>15}  {pl['khoan_muc_ma']:14} {pl['cong_trinh_ma']}")
            return
        print({k: v for k, v in tk.items() if k != "canh_bao"})
        for c in tk["canh_bao"]:
            print("  ! " + c)
    elif a.cmd == "doi-chieu":
        from .reconcile import compare
        if a.misa:
            nguon = a.misa
        else:
            from . import misa_db
            tu, den = utils.parse_date(a.tu), utils.parse_date(a.den)
            if not tu or not den:
                raise SystemExit("Đọc thẳng CSDL MISA cần --tu và --den.")
            nguon = misa_db.bang_ke_mua_vao(tu, den)
            print(f"Đọc {len(nguon)} dòng hóa đơn mua vào từ CSDL MISA ({a.tu} - {a.den}).")
        kq = compare.doi_chieu(nguon, a.thue)
        print(kq["Tong hop"].to_string(index=False))
        out = compare.xuat_bao_cao(kq, utils.DATA_DIR / "bao_cao" / f"DoiChieu_{datetime.now():%Y%m%d_%H%M}.xlsx")
        print(f"\n=> Báo cáo: {out}")
    elif a.cmd == "misa":
        kham_pha_misa()
    elif a.cmd == "luong":
        from .payroll import engine
        df = engine.tinh_luong(a.nhan_su, a.cham_cong, a.danh_gia)
        print(df[["Mã NV", "Họ tên", "Ngày công", "Tổng thu nhập", "Tổng BH NLĐ", "Thuế TNCN", "THỰC LĨNH"]].to_string(index=False))
        out = engine.xuat_bang_luong(df, utils.DATA_DIR / "bao_cao" / f"BangLuong_{a.thang.replace('/', '_')}.xlsx", a.thang)
        print(f"\n=> Bảng lương: {out}")


def kham_pha_misa():
    """Dò SQL Server của MISA, liệt kê dữ liệu kế toán và bảng hóa đơn; ghi ra misa_kham_pha.txt để gửi người cài đặt."""
    import os

    import pandas as pd

    from . import misa_db

    out = []

    def log(s=""):
        print(s)
        out.append(str(s))

    try:
        sv = utils.env("KT_MISA_SERVER") or misa_db.tu_do_server()
        log(f"SQL Server của MISA: {sv}")
        os.environ["KT_MISA_SERVER"] = sv
        dbs = misa_db.danh_sach_du_lieu(sv)
        log("Các dữ liệu kế toán (CSDL): " + ", ".join(dbs))
        db = utils.env("KT_MISA_DB") or (dbs[0] if len(dbs) == 1 else "")
        if not db:
            log("\n=> Đặt KT_MISA_DB=<tên dữ liệu đang dùng> trong .env rồi chạy lại để dò bảng hóa đơn.")
        else:
            os.environ["KT_MISA_DB"] = db
            conn = misa_db.ket_noi(db, sv)
            ung_vien = misa_db.kham_pha(conn)
            log(f"\nBảng có cột số hóa đơn trong {db}:")
            with pd.option_context("display.max_colwidth", 200, "display.width", 250):
                log(ung_vien.to_string(index=False) if not ung_vien.empty else "(không thấy)")
            for b in ung_vien.head(6)["bang"]:
                log(f"\n--- 3 dòng mẫu của {b} (chỉ tên cột + kiểu, không chép số liệu) ---")
                mau = misa_db.mau_du_lieu(b, 3, conn)
                log(", ".join(f"{c}:{type(v).__name__}" for c, v in (mau.iloc[0].items() if len(mau) else [])))
            conn.close()
    except Exception as ex:  # noqa: BLE001
        log(f"LỖI: {ex}")
    p = utils.ROOT / "misa_kham_pha.txt"
    p.write_text("\n".join(out), encoding="utf-8")
    print(f"\n=> Đã ghi {p}. Gửi file này cho người cài đặt.")


if __name__ == "__main__":
    sys.exit(main())
