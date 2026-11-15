"""Luồng xử lý hóa đơn: email/thư mục -> đọc -> phân loại -> kho chờ duyệt -> file nhập MISA."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .. import utils
from . import extractor, misa_export, store
from .gmail_reader import EmailHoaDon


def chay_tu_email(so_ngay: int | None = None, dung_link: bool = True, gan_nhan: bool = True) -> dict:
    from .gmail_reader import GmailIMAP

    cfg = utils.settings()["hoa_don"]
    g = GmailIMAP()
    emails = g.lay_email(so_ngay or cfg.get("so_ngay_gan_nhat", 7), cfg.get("chi_chua_doc", False), cfg.get("thu_muc_email", "INBOX"))
    tk = {"email": len(emails), "moi": 0, "trung": 0, "canh_bao": []}
    for e in emails:
        if store.email_da_xu_ly(e.uid):
            continue
        ket_qua, cb = extractor.trich_xuat_tu_email(e, dung_link=dung_link)
        tk["canh_bao"].extend(cb)
        for hd, pl in ket_qua:
            if store.luu(hd.model_dump(), pl.model_dump(), e.uid, e.subject):
                tk["moi"] += 1
            else:
                tk["trung"] += 1
        if ket_qua:
            store.danh_dau_email(e.uid)
            if gan_nhan:
                try:
                    g.gan_nhan(e.uid, cfg.get("nhan_da_xu_ly", "AI_Da_Nhap"))
                except Exception as ex:  # noqa: BLE001
                    tk["canh_bao"].append(f"Không gắn nhãn được email {e.uid}: {ex}")
    g.dong()
    return tk


def chay_tu_thu_muc(thu_muc: str | Path, ghi_chu: str = "") -> dict:
    """Xử lý các file XML/PDF/ZIP/ảnh đã tải sẵn vào một thư mục (không cần email)."""
    thu_muc = Path(thu_muc)
    tk = {"file": 0, "moi": 0, "trung": 0, "canh_bao": []}
    for f in sorted(thu_muc.iterdir()):
        if not f.is_file() or f.suffix.lower() not in (".xml", ".pdf", ".zip", ".png", ".jpg", ".jpeg"):
            continue
        tk["file"] += 1
        e = EmailHoaDon(uid=f"file:{f.name}", subject=f.name, nguoi_gui="", ngay="", ghi_chu=ghi_chu, body_text="",
                        attachments=[(f.name, f.read_bytes())], links=[])
        ket_qua, cb = extractor.trich_xuat_tu_email(e, dung_link=False)
        tk["canh_bao"].extend(cb)
        for hd, pl in ket_qua:
            if store.luu(hd.model_dump(), pl.model_dump(), e.uid, e.subject):
                tk["moi"] += 1
            else:
                tk["trung"] += 1
    return tk


def xuat_misa(trang_thai: str = "cho_duyet", out_dir: str | Path | None = None, danh_dau_da_xuat: bool = True) -> tuple[Path | None, int]:
    items = store.danh_sach(trang_thai)
    if not items:
        return None, 0
    out_dir = Path(out_dir) if out_dir else utils.DATA_DIR / "misa_import"
    out = misa_export.xuat_file_mua_hang(items, out_dir / f"MISA_MuaHang_{datetime.now():%Y%m%d_%H%M}.xlsx")
    if danh_dau_da_xuat:
        for it in items:
            store.cap_nhat(it["khoa"], trang_thai="da_xuat")
    return out, len(items)
