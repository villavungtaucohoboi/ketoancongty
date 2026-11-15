"""Kho lưu hóa đơn đã xử lý (SQLite) - chống nhập trùng và lưu trạng thái duyệt."""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path

from .. import utils

DB = utils.DATA_DIR / "ketoan.sqlite"


def _conn():
    utils.ensure_dir(DB.parent)
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS hoa_don (
        khoa TEXT PRIMARY KEY, mst_ban TEXT, ky_hieu TEXT, so_hd TEXT, ngay_hd TEXT, ten_ban TEXT,
        tong REAL, khoan_muc TEXT, cong_trinh TEXT, nguon TEXT, email_uid TEXT, email_subject TEXT,
        trang_thai TEXT DEFAULT 'cho_duyet', json TEXT, tao_luc TEXT)""")
    c.execute("CREATE TABLE IF NOT EXISTS email_da_xu_ly (uid TEXT PRIMARY KEY, luc TEXT)")
    return c


def khoa_hoa_don(mst_ban: str, ky_hieu: str, so_hd: str) -> str:
    return f"{mst_ban}|{(ky_hieu or '').upper()}|{(so_hd or '').lstrip('0')}"


def da_co(khoa: str) -> bool:
    with _conn() as c:
        return c.execute("SELECT 1 FROM hoa_don WHERE khoa=?", (khoa,)).fetchone() is not None


def luu(hd: dict, phan_loai: dict, email_uid: str = "", email_subject: str = "") -> bool:
    khoa = khoa_hoa_don(hd.get("mst_nguoi_ban", ""), hd.get("ky_hieu", ""), hd.get("so_hoa_don", ""))
    if da_co(khoa):
        return False
    with _conn() as c:
        c.execute("INSERT INTO hoa_don VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
            khoa, hd.get("mst_nguoi_ban"), hd.get("ky_hieu"), hd.get("so_hoa_don"), hd.get("ngay_hoa_don"), hd.get("ten_nguoi_ban"),
            hd.get("tong_thanh_toan"), phan_loai.get("khoan_muc_ma"), phan_loai.get("cong_trinh_ma"), hd.get("nguon"),
            email_uid, email_subject, "cho_duyet", json.dumps({"hd": hd, "pl": phan_loai}, ensure_ascii=False), datetime.now().isoformat()))
    return True


def email_da_xu_ly(uid: str) -> bool:
    with _conn() as c:
        return c.execute("SELECT 1 FROM email_da_xu_ly WHERE uid=?", (uid,)).fetchone() is not None


def danh_dau_email(uid: str):
    with _conn() as c:
        c.execute("INSERT OR IGNORE INTO email_da_xu_ly VALUES (?,?)", (uid, datetime.now().isoformat()))


def danh_sach(trang_thai: str | None = None) -> list[dict]:
    with _conn() as c:
        q = "SELECT json, trang_thai, khoa FROM hoa_don" + (" WHERE trang_thai=?" if trang_thai else "") + " ORDER BY tao_luc DESC"
        rows = c.execute(q, (trang_thai,) if trang_thai else ()).fetchall()
    out = []
    for js, tt, khoa in rows:
        d = json.loads(js)
        d["trang_thai"] = tt
        d["khoa"] = khoa
        out.append(d)
    return out


def cap_nhat(khoa: str, trang_thai: str | None = None, phan_loai: dict | None = None):
    with _conn() as c:
        if phan_loai is not None:
            row = c.execute("SELECT json FROM hoa_don WHERE khoa=?", (khoa,)).fetchone()
            if row:
                d = json.loads(row[0])
                d["pl"].update(phan_loai)
                c.execute("UPDATE hoa_don SET json=?, khoan_muc=?, cong_trinh=? WHERE khoa=?",
                          (json.dumps(d, ensure_ascii=False), d["pl"].get("khoan_muc_ma"), d["pl"].get("cong_trinh_ma"), khoa))
        if trang_thai:
            c.execute("UPDATE hoa_don SET trang_thai=? WHERE khoa=?", (trang_thai, khoa))
