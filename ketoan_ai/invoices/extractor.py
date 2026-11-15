"""Từ một email (hoặc một file) -> danh sách HoaDon đã đọc + phân loại khoản mục / công trình."""
from __future__ import annotations

import io
import zipfile
from pathlib import Path

from .. import ai, utils
from . import xml_parser
from .gmail_reader import EmailHoaDon


def _mo_zip(data: bytes) -> list[tuple[str, bytes]]:
    out = []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for n in z.namelist():
                if n.lower().endswith((".xml", ".pdf", ".png", ".jpg", ".jpeg")) and not n.endswith("/"):
                    out.append((Path(n).name, z.read(n)))
    except zipfile.BadZipFile:
        pass
    return out


def doc_file(ten: str, data: bytes) -> ai.HoaDon | None:
    """Đọc một file đính kèm. Ưu tiên XML (chính xác, miễn phí) > PDF > ảnh."""
    low = ten.lower()
    if low.endswith(".xml") or (data[:5] == b"<?xml" and xml_parser.la_hoa_don_xml(data)):
        if xml_parser.la_hoa_don_xml(data):
            return xml_parser.parse_xml(data)
        return None
    if low.endswith(".pdf") or data[:4] == b"%PDF":
        return ai.doc_hoa_don_tu_file(data, "application/pdf")
    if low.endswith((".png", ".jpg", ".jpeg")):
        return ai.doc_hoa_don_tu_file(data, "image/png" if low.endswith(".png") else "image/jpeg")
    return None


def _la_hoa_don_mua_vao(hd: ai.HoaDon) -> bool:
    mst_cty = utils.settings()["cong_ty"].get("mst", "")
    return not mst_cty or not hd.mst_nguoi_mua or hd.mst_nguoi_mua == mst_cty


def trich_xuat_tu_email(e: EmailHoaDon, dung_link: bool = True) -> tuple[list[tuple[ai.HoaDon, ai.PhanLoaiHoaDon]], list[str]]:
    """Trả về ([(HoaDon, PhanLoai)], [cảnh báo])."""
    canh_bao: list[str] = []
    files: list[tuple[str, bytes]] = []
    for ten, data in e.attachments:
        if ten.lower().endswith(".zip"):
            files.extend(_mo_zip(data))
        else:
            files.append((ten, data))
    # Nếu có XML thì bỏ PDF cùng tên/cùng hóa đơn (tránh đọc 2 lần)
    xmls = [f for f in files if f[0].lower().endswith(".xml")]
    if xmls:
        files = xmls
    ket_qua: list[ai.HoaDon] = []
    for ten, data in files:
        try:
            hd = doc_file(ten, data)
            if hd:
                ket_qua.append(hd)
        except Exception as ex:  # noqa: BLE001
            canh_bao.append(f"{e.subject[:60]} / {ten}: {ex}")
    if not ket_qua and e.links and dung_link:
        from . import link_fetcher
        for url in e.links[:2]:
            try:
                data, mt = link_fetcher.tai_tu_link(url)
                if mt == "text/xml":
                    hd = xml_parser.parse_xml(data)
                else:
                    hd = ai.doc_hoa_don_tu_file(data, mt)
                    hd.nguon = "link"
                ket_qua.append(hd)
                break
            except Exception as ex:  # noqa: BLE001
                canh_bao.append(f"{e.subject[:60]} / link {url[:50]}...: {ex}")
    if not ket_qua:
        canh_bao.append(f"{e.subject[:80]}: không tìm thấy hóa đơn (không có XML/PDF và không đọc được link).")
    out = []
    for hd in ket_qua:
        if not _la_hoa_don_mua_vao(hd):
            canh_bao.append(f"HĐ {hd.ky_hieu}/{hd.so_hoa_don}: MST người mua {hd.mst_nguoi_mua} khác MST công ty - bỏ qua (có thể là HĐ đầu ra).")
            continue
        # loại bỏ dòng hàng là ghi chú không có tiền
        hd.dong_hang = [d for d in hd.dong_hang if d.thanh_tien or d.so_luong]
        pl = ai.phan_loai_hoa_don(hd, e.ghi_chu)
        out.append((hd, pl))
    return out, canh_bao
