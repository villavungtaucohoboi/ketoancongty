"""Đọc trực tiếp (CHỈ ĐỌC) cơ sở dữ liệu MISA SME.NET trên máy chủ nội bộ của công ty.

MISA SME.NET lưu dữ liệu trong Microsoft SQL Server cài kèm (instance dạng .\\MISASME2023).
Module này chỉ chạy câu lệnh SELECT, không bao giờ ghi/sửa/xóa dữ liệu MISA.
Dữ liệu đọc ra chỉ dùng trong máy nội bộ để đối chiếu, không gửi lên AI hay Internet.

Cấu hình trong .env:
    KT_MISA_SERVER=.\\MISASME2023      (máy chủ\\instance SQL của MISA; bỏ trống = tự dò)
    KT_MISA_DB=TEN_DU_LIEU_KE_TOAN       (tên CSDL = tên dữ liệu kế toán đang mở trong MISA)
    KT_MISA_USER= / KT_MISA_PASSWORD=   (bỏ trống = đăng nhập Windows của máy đang chạy app)

Câu truy vấn lấy bảng kê mua vào nằm trong config/misa_sql.yaml để sửa không cần đụng code.
Chạy `python -m ketoan_ai.cli misa kham-pha` để dò tên bảng/cột thật trong CSDL của công ty.
"""
from __future__ import annotations

import re
from datetime import date
from typing import Any, Callable

import pandas as pd

from . import utils

# Thứ tự ưu tiên driver ODBC trên Windows. "SQL Server" luôn có sẵn trên mọi máy Windows.
DRIVERS = ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "SQL Server Native Client 11.0", "SQL Server"]
NAM_MISA = range(2026, 2014, -1)  # các phiên bản MISASME2026 ... MISASME2015

_connect_factory: Callable[[], Any] | None = None  # dùng cho kiểm thử


def cau_hinh() -> dict:
    return utils.load_yaml("misa_sql.yaml")


def da_cau_hinh() -> bool:
    return bool(utils.env("KT_MISA_DB")) or _connect_factory is not None


def _driver() -> str:
    import pyodbc

    co = set(pyodbc.drivers())
    for d in DRIVERS:
        if d in co:
            return d
    raise RuntimeError("Không tìm thấy driver ODBC cho SQL Server trên máy này. Cài 'ODBC Driver 18 for SQL Server' của Microsoft.")


def _conn_str(server: str, database: str = "master") -> str:
    s = f"DRIVER={{{_driver()}}};SERVER={server};DATABASE={database};"
    user = utils.env("KT_MISA_USER")
    if user:
        s += f"UID={user};PWD={utils.env('KT_MISA_PASSWORD')};"
    else:
        s += "Trusted_Connection=yes;"
    if "18" in _driver():
        s += "TrustServerCertificate=yes;"
    return s + "ApplicationIntent=ReadOnly;"


def ket_noi(database: str | None = None, server: str | None = None):
    if _connect_factory is not None:
        return _connect_factory()
    import pyodbc

    server = server or utils.env("KT_MISA_SERVER") or tu_do_server()
    database = database or utils.env("KT_MISA_DB")
    if not database:
        raise RuntimeError("Chưa đặt KT_MISA_DB (tên dữ liệu kế toán). Chạy 'kham_pha_misa.bat' để xem danh sách.")
    conn = pyodbc.connect(_conn_str(server, database), timeout=8, autocommit=True)
    return conn


def tu_do_server() -> str:
    """Thử các instance SQL Server mà MISA hay cài, trả về instance đầu tiên kết nối được."""
    import pyodbc

    ung_vien = [utils.env("KT_MISA_SERVER")] if utils.env("KT_MISA_SERVER") else []
    ung_vien += [rf".\MISASME{y}" for y in NAM_MISA] + [r".\SQLEXPRESS", "."]
    loi = []
    for sv in ung_vien:
        try:
            pyodbc.connect(_conn_str(sv), timeout=3).close()
            return sv
        except Exception as ex:  # noqa: BLE001
            loi.append(f"{sv}: {str(ex)[:80]}")
    raise RuntimeError("Không kết nối được SQL Server của MISA. Đã thử:\n" + "\n".join(loi[:6]))


def _select(conn, sql: str, params: tuple = ()) -> pd.DataFrame:
    if not re.match(r"^\s*(--[^\n]*\n\s*)*(SELECT|WITH)\b", sql, re.I) or re.search(r"\b(INSERT|UPDATE|DELETE|DROP|ALTER|TRUNCATE|EXEC|MERGE|CREATE|GRANT)\b", sql, re.I):
        raise ValueError("Chỉ cho phép câu lệnh SELECT khi đọc CSDL MISA.")
    cur = conn.cursor()
    cur.execute(sql, params)
    cols = [c[0] for c in cur.description]
    return pd.DataFrame.from_records([tuple(r) for r in cur.fetchall()], columns=cols)


# ---------------------------------------------------------------------------
# 1. Dò cấu trúc (chạy 1 lần trên máy chủ MISA, gửi kết quả cho người cài đặt)
# ---------------------------------------------------------------------------
def danh_sach_du_lieu(server: str | None = None) -> list[str]:
    """Liệt kê các CSDL (mỗi 'dữ liệu kế toán' của MISA là một CSDL)."""
    import pyodbc

    server = server or utils.env("KT_MISA_SERVER") or tu_do_server()
    conn = pyodbc.connect(_conn_str(server), timeout=8, autocommit=True)
    df = _select(conn, "SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name")
    conn.close()
    return df["name"].tolist()


def kham_pha(conn=None) -> pd.DataFrame:
    """Tìm các bảng có cột số hóa đơn + mã số thuế: ứng viên cho bảng kê hóa đơn mua vào."""
    own = conn is None
    conn = conn or ket_noi()
    sql = """
SELECT t.TABLE_NAME AS bang, c.COLUMN_NAME AS cot, c.DATA_TYPE AS kieu
FROM INFORMATION_SCHEMA.COLUMNS c JOIN INFORMATION_SCHEMA.TABLES t
  ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_SCHEMA = c.TABLE_SCHEMA
WHERE t.TABLE_TYPE IN ('BASE TABLE','VIEW')
  AND (c.COLUMN_NAME LIKE '%Inv%No%' OR c.COLUMN_NAME LIKE '%Invoice%' OR c.COLUMN_NAME LIKE '%Series%'
       OR c.COLUMN_NAME LIKE '%TaxCode%' OR c.COLUMN_NAME LIKE '%VATAmount%' OR c.COLUMN_NAME LIKE '%InvDate%')
ORDER BY t.TABLE_NAME, c.COLUMN_NAME"""
    df = _select(conn, sql)
    if own:
        conn.close()
    if df.empty:
        return df
    g = df.groupby("bang")["cot"].apply(list).reset_index()
    g["co_so_hd"] = g["cot"].map(lambda cs: any(re.search(r"inv.*no|invoiceno", c, re.I) for c in cs))
    g["co_mst"] = g["cot"].map(lambda cs: any(re.search(r"taxcode", c, re.I) for c in cs))
    g = g[g["co_so_hd"]].sort_values(["co_mst"], ascending=False)
    g["cot"] = g["cot"].map(lambda cs: ", ".join(cs))
    return g[["bang", "co_mst", "cot"]]


def mau_du_lieu(bang: str, n: int = 3, conn=None) -> pd.DataFrame:
    if not re.fullmatch(r"[A-Za-z0-9_]+", bang):
        raise ValueError("Tên bảng không hợp lệ.")
    own = conn is None
    conn = conn or ket_noi()
    df = _select(conn, f"SELECT TOP {int(n)} * FROM [{bang}]")
    if own:
        conn.close()
    return df


# ---------------------------------------------------------------------------
# 2. Bảng kê hóa đơn mua vào trong kỳ (dùng cho đối chiếu)
# ---------------------------------------------------------------------------
def bang_ke_mua_vao(tu: date, den: date, conn=None) -> pd.DataFrame:
    """Trả về DataFrame cột chuẩn: mst, so_hd, ky_hieu, mau_so, ngay, ten, tien_hang, tien_thue."""
    cfg = cau_hinh()
    sql = cfg["bang_ke_mua_vao"]["sql"]
    own = conn is None
    conn = conn or ket_noi()
    df = _select(conn, sql, (tu.isoformat(), den.isoformat()))
    if own:
        conn.close()
    doi_ten = {v: k for k, v in cfg["bang_ke_mua_vao"].get("cot", {}).items()}
    df = df.rename(columns=doi_ten)
    thieu = [c for c in ("mst", "so_hd", "tien_hang", "tien_thue") if c not in df]
    if thieu:
        raise ValueError(f"Câu truy vấn MISA thiếu cột {thieu}. Sửa config/misa_sql.yaml.")
    if "ngay" in df:
        df["ngay"] = df["ngay"].map(utils.parse_date)
    return df
