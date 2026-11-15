"""Đọc hóa đơn điện tử dạng XML theo chuẩn Thông tư 78/2021 (định dạng của Tổng cục Thuế).

XML là nguồn CHÍNH XÁC NHẤT (không cần AI). Hầu hết email hóa đơn đính kèm file .xml hoặc .zip chứa .xml.
Bộ đọc dùng tìm theo tên thẻ (bỏ namespace) nên chịu được khác biệt nhỏ giữa các nhà cung cấp HĐĐT.
"""
from __future__ import annotations

import re
from lxml import etree

from .. import utils
from ..ai import DongHang, HoaDon


def _text(node, *paths: str) -> str:
    for p in paths:
        r = node.xpath(f".//*[local-name()='{p}']")
        if r and r[0].text:
            return r[0].text.strip()
    return ""


def _num(node, *paths: str) -> float:
    return utils.parse_money(_text(node, *paths))


def la_hoa_don_xml(data: bytes) -> bool:
    head = data[:4000].decode("utf-8", "ignore")
    return "<HDon" in head or "<DLHDon" in head or "<TTChung" in head or "<NDHDon" in head or "invoice" in head.lower()


def parse_xml(data: bytes) -> HoaDon:
    parser = etree.XMLParser(recover=True, huge_tree=True, resolve_entities=False, no_network=True)
    root = etree.fromstring(data, parser=parser)
    if root is None:
        raise ValueError("Không đọc được XML hóa đơn.")
    hd_nodes = root.xpath("//*[local-name()='DLHDon']") or [root]
    dl = hd_nodes[0]
    nban = (dl.xpath(".//*[local-name()='NBan']") or [dl])[0]
    nmua = (dl.xpath(".//*[local-name()='NMua']") or [dl])[0]
    ttoan = (dl.xpath(".//*[local-name()='TToan']") or [dl])[0]

    ngay = _text(dl, "NLap", "NgayLap", "NgayHD")
    d = utils.parse_date(ngay)
    dong = []
    for i, h in enumerate(dl.xpath(".//*[local-name()='HHDVu']"), start=1):
        ts_raw = _text(h, "TSuat", "ThueSuat")
        ts = None
        m = re.search(r"(\d+(?:[.,]\d+)?)", ts_raw)
        if m:
            ts = float(m.group(1).replace(",", "."))
        elif "KCT" in ts_raw.upper() or "KKKNT" in ts_raw.upper():
            ts = None
        thanh_tien = _num(h, "ThTien", "ThanhTien")
        tien_thue = _num(h, "TThue", "TienThue")
        if not tien_thue and ts:
            tien_thue = round(thanh_tien * ts / 100)
        dong.append(DongHang(stt=i, ten_hang=_text(h, "THHDVu", "TenHang", "TenHHDV"), dvt=_text(h, "DVTinh", "DonViTinh"),
                             so_luong=_num(h, "SLuong", "SoLuong"), don_gia=_num(h, "DGia", "DonGia"),
                             thanh_tien=thanh_tien, thue_suat=ts, tien_thue=tien_thue))
    tien_hang = _num(ttoan, "TgTCThue", "TongTienChuaThue") or sum(x.thanh_tien for x in dong)
    tien_thue = _num(ttoan, "TgTThue", "TongTienThue") or sum(x.tien_thue for x in dong)
    tong = _num(ttoan, "TgTTTBSo", "TongTienThanhToan") or (tien_hang + tien_thue)
    return HoaDon(
        mau_so=_text(dl, "KHMSHDon", "MauSo"), ky_hieu=_text(dl, "KHHDon", "KyHieu"), so_hoa_don=_text(dl, "SHDon", "SoHD", "SoHoaDon"),
        ngay_hoa_don=utils.fmt_date(d) if d else ngay,
        ten_nguoi_ban=_text(nban, "Ten", "TenDonVi"), mst_nguoi_ban=re.sub(r"\D", "", _text(nban, "MST", "MaSoThue")), dia_chi_nguoi_ban=_text(nban, "DChi", "DiaChi"),
        ten_nguoi_mua=_text(nmua, "Ten", "TenDonVi"), mst_nguoi_mua=re.sub(r"\D", "", _text(nmua, "MST", "MaSoThue")),
        ma_cqt=_text(root, "MCCQT", "MaCQT"), tien_hang=tien_hang, tien_thue=tien_thue, tong_thanh_toan=tong,
        hinh_thuc_thanh_toan=_text(dl, "HTTToan", "HinhThucThanhToan"), dong_hang=dong, nguon="xml", do_tin_cay=1.0,
    )
