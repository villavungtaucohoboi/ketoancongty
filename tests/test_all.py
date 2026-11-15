"""Kiểm thử offline (không cần API key) cho 4 phân hệ. Chạy: python -m pytest -q"""
import os
from pathlib import Path

import pytest

from conftest import SAMPLES
from ketoan_ai import ai, utils
from ketoan_ai.bank import misa_export as bank_misa, parser
from ketoan_ai.invoices import extractor, misa_export as inv_misa, store, xml_parser
from ketoan_ai.invoices.gmail_reader import parse_message
from ketoan_ai.payroll import engine
from ketoan_ai.reconcile import compare


@pytest.fixture(autouse=True)
def no_api_key(monkeypatch, tmp_path):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.delenv("ANTHROPIC_AUTH_TOKEN", raising=False)
    monkeypatch.setattr(store, "DB", tmp_path / "t.sqlite")


def test_parse_money_vn():
    assert utils.parse_money("11.340.000") == 11340000
    assert utils.parse_money("3,456,000 VND") == 3456000
    assert utils.parse_money("1.234.567,89") == pytest.approx(1234567.89)
    assert utils.parse_money(float("nan")) == 0
    assert utils.parse_money("") == 0


def test_bo_dau_va_match():
    assert utils.bo_dau("Bê tông Nam Hà Nội") == "BE TONG NAM HA NOI"
    assert utils.match_cong_trinh("HCIC TT CONG NO BE TONG CT NAM HA NOI")["ma"] == "25_TC_NAM HA NOI"
    assert utils.match_khoan_muc("Bê tông thương phẩm M250")["ma"] == "NVL_Betong"
    assert utils.match_doi_tuong(mst="0104763486")["ma"] == "2018NB33"


def test_sao_ke_parser_va_xuat_misa(tmp_path):
    gds = parser.doc_sao_ke_tu_file(SAMPLES / "sao_ke_ACB_mau.xlsx")
    assert len(gds) == 12
    assert gds[0]["tien_ra"] == 11340000 and gds[0]["tien_vao"] == 0
    assert gds[1]["tien_vao"] == 32417280
    pl = ai.phan_loai_giao_dich(gds)
    assert [p.loai for p in pl][:3] == ["chi", "thu", "chi"]
    phi = next(p for g, p in zip(gds, pl) if "PHI QUAN LY" in g["noi_dung"])
    assert phi.tk_doi_ung == "6428" and phi.do_tin_cay >= 0.9
    lai = next(p for g, p in zip(gds, pl) if "LAI TIEN GUI" in g["noi_dung"])
    assert lai.loai == "thu" and lai.tk_doi_ung == "515"
    be_tong = next(p for g, p in zip(gds, pl) if "BE TONG" in g["noi_dung"])
    assert be_tong.cong_trinh_ma == "25_TC_NAM HA NOI" and be_tong.khoan_muc_ma == "NVL_Betong"
    out = bank_misa.xuat_file_misa(gds, pl, utils.settings()["sao_ke"]["tabs"]["ACB_CTY"], tmp_path / "x.xlsx")
    from openpyxl import load_workbook
    wb = load_workbook(out)
    assert wb.sheetnames == ["Thu tien gui", "Chi tien gui"]
    assert wb["Chi tien gui"].max_row - 1 == 9 and wb["Thu tien gui"].max_row - 1 == 3
    row = list(wb["Chi tien gui"].iter_rows(min_row=2, max_row=2, values_only=True))[0]
    assert row[8] == "331" and row[9] == "1121" and row[10] == 11340000


def test_xml_hoa_don():
    hd = xml_parser.parse_xml((SAMPLES / "hoa_don_84_mau.xml").read_bytes())
    assert hd.so_hoa_don == "84" and hd.ky_hieu == "C26TMT" and hd.ngay_hoa_don == "29/10/2026"
    assert hd.mst_nguoi_ban == "0108640735" and hd.mst_nguoi_mua == "0103775584"
    assert hd.tien_hang == 3200000 and hd.tien_thue == 256000 and hd.tong_thanh_toan == 3456000
    assert hd.dong_hang[0].ten_hang == "Máy laser 5 tia LSG686SPD" and hd.dong_hang[0].thue_suat == 8
    assert hd.ma_cqt == "000C5D0F05132843FB959160DF96CC7EAF"


def test_email_forward_zip_va_ghi_chu_cong_trinh():
    e = parse_message("1", (SAMPLES / "email_hoa_don_mau.eml").read_bytes())
    assert e.ghi_chu == "Công trình Nam Hà Nội"
    assert e.attachments[0][0].endswith(".zip")
    assert e.links == ["https://tracuu.easyinvoice.com.vn/?code=A9HL5FRU2"]
    kq, cb = extractor.trich_xuat_tu_email(e, dung_link=False)
    assert not cb and len(kq) == 1
    hd, pl = kq[0]
    assert hd.so_hoa_don == "604" and pl.khoan_muc_ma == "NVL_Betong" and pl.cong_trinh_ma == "25_TC_NAM HA NOI"


def test_store_chong_trung_va_xuat_mua_hang(tmp_path):
    hd = xml_parser.parse_xml((SAMPLES / "hoa_don_84_mau.xml").read_bytes())
    pl = ai.phan_loai_hoa_don(hd, "Nhà B An Khánh")
    assert pl.cong_trinh_ma == "25TC_ANKHANH_NHA B"
    assert store.luu(hd.model_dump(), pl.model_dump(), "u1", "s") is True
    assert store.luu(hd.model_dump(), pl.model_dump(), "u2", "s") is False  # trùng
    items = store.danh_sach("cho_duyet")
    assert len(items) == 1
    store.cap_nhat(items[0]["khoa"], phan_loai={"khoan_muc_ma": "Chi phí NVL Phụ"})
    items = store.danh_sach("cho_duyet")
    assert items[0]["pl"]["khoan_muc_ma"] == "Chi phí NVL Phụ"
    out = inv_misa.xuat_file_mua_hang(items, tmp_path / "mh.xlsx")
    from openpyxl import load_workbook
    ws = load_workbook(out).active
    r = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    assert r[3] == "NCC_640735" and r[9] == "C26TMT" and r[10] == "84" and r[17] == 3200000 and r[19] == 256000
    assert r[23] == "Chi phí NVL Phụ" and r[24] == "Thi công công trình Nhà B An Khánh"


def test_doi_chieu(tmp_path):
    kq = compare.doi_chieu(SAMPLES / "bang_ke_misa_mau.xlsx", SAMPLES / "bang_ke_thue_mau.xlsx")
    th = kq["Tong hop"].set_index("Chỉ tiêu")["Số lượng"]
    assert th["Khớp hoàn toàn"] == 3
    assert th["Chỉ có trên Thuế (chưa nhập MISA)"] == 2
    assert th["Chỉ có trên MISA (thuế chưa có / nhập sai)"] == 1
    assert set(kq["Chi co tren Thue"]["Số HĐ"]) == {"603", "156"}
    assert list(kq["Chi co tren MISA"]["Số HĐ"]) == ["3561"]
    out = compare.xuat_bao_cao(kq, tmp_path / "dc.xlsx")
    assert out.exists()


def test_thue_tncn_luy_tien():
    bieu = utils.load_yaml("luong_rules.yaml")["bieu_thue"]
    assert engine.thue_tncn(0, bieu) == 0
    assert engine.thue_tncn(5_000_000, bieu) == 250_000
    assert engine.thue_tncn(10_000_000, bieu) == 750_000
    assert engine.thue_tncn(20_000_000, bieu) == 250_000 + 500_000 + 1_200_000 + 400_000


def test_tinh_luong(tmp_path):
    df = engine.tinh_luong(SAMPLES / "nhan_su_mau.xlsx", SAMPLES / "cham_cong_mau.xlsx", SAMPLES / "danh_gia_mau.xlsx")
    assert len(df) == 3
    b = df.set_index("Mã NV").loc["NV002"]
    assert b["Ngày công"] == 26 and b["Lương theo công"] == 15_000_000
    assert b["Thưởng KPI"] == 750_000 and b["Phạt/trừ"] == 100_000
    assert b["Tổng thu nhập"] == 17_580_000
    assert b["BHXH (8%)"] == 800_000 and b["BHYT (1.5%)"] == 150_000 and b["BHTN (1%)"] == 100_000
    assert b["Thuế TNCN"] == 5_000 and b["THỰC LĨNH"] == 16_525_000
    out = engine.xuat_bang_luong(df, tmp_path / "bl.xlsx", "10/2026")
    assert out.exists()
