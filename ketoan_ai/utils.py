"""Hàm tiện ích dùng chung: đọc cấu hình, chuẩn hoá tiếng Việt, số tiền, ngày tháng."""
from __future__ import annotations

import os
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT / "config"
DATA_DIR = ROOT / "data"

load_dotenv(ROOT / ".env")


def load_yaml(name: str) -> dict:
    path = CONFIG_DIR / name
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def settings() -> dict:
    return load_yaml("settings.yaml")


def khoan_muc_list() -> list[dict]:
    return load_yaml("khoan_muc_chi_phi.yaml").get("khoan_muc", [])


def cong_trinh_list() -> list[dict]:
    return load_yaml("cong_trinh.yaml").get("cong_trinh", [])


def doi_tuong_list() -> list[dict]:
    return load_yaml("doi_tuong.yaml").get("doi_tuong", [])


def misa_templates() -> dict:
    return load_yaml("misa_templates.yaml")


def bo_dau(s: str | None) -> str:
    """Bỏ dấu tiếng Việt, đưa về chữ hoa, gộp khoảng trắng: 'Bê tông' -> 'BE TONG'."""
    if not s:
        return ""
    s = str(s).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().upper()


def parse_money(v: Any) -> float:
    """Đọc số tiền kiểu Việt Nam: '11.340.000', '3,456,000 VND', '1.234.567,89' -> float."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if v != v else float(v)  # NaN -> 0
    s = str(v).strip()
    if not s:
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s or s in {"-", ".", ","}:
        return 0.0
    # Nếu có cả '.' và ',' thì ký tự xuất hiện sau cùng là dấu thập phân
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        # '1,234,567' (nghìn) hay '1234,5' (thập phân)?
        parts = s.split(",")
        if all(len(p) == 3 for p in parts[1:]):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if all(len(p) == 3 for p in parts[1:]):
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


_DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d.%m.%Y", "%Y-%m-%dT%H:%M:%S", "%d/%m/%Y %H:%M:%S"]


def parse_date(v: Any) -> date | None:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s[: len(fmt) + 4], fmt).date()
        except ValueError:
            continue
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def fmt_date(d: date | None) -> str:
    return d.strftime("%d/%m/%Y") if d else ""


def fmt_money(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".")


def find_keyword(text: str, keywords: list[str]) -> bool:
    t = bo_dau(text)
    return any(bo_dau(k) in t for k in keywords if k)


def match_cong_trinh(text: str) -> dict | None:
    """Tìm công trình theo bí danh trong một đoạn text (không dấu, không phân biệt hoa thường)."""
    t = bo_dau(text)
    best = None
    for ct in cong_trinh_list():
        for alias in [ct.get("ten", "")] + list(ct.get("bi_danh", [])):
            a = bo_dau(alias)
            if a and a in t and (best is None or len(a) > best[0]):
                best = (len(a), ct)
    return best[1] if best else None


def match_khoan_muc(text: str) -> dict | None:
    t = bo_dau(text)
    best = None
    for km in khoan_muc_list():
        for kw in km.get("tu_khoa", []):
            k = bo_dau(kw)
            if k and k in t and (best is None or len(k) > best[0]):
                best = (len(k), km)
    return best[1] if best else None


def match_doi_tuong(text: str = "", mst: str = "") -> dict | None:
    mst = re.sub(r"\D", "", mst or "")
    for dt in doi_tuong_list():
        if mst and re.sub(r"\D", "", str(dt.get("mst") or "")) == mst:
            return dt
    t = bo_dau(text)
    for dt in doi_tuong_list():
        for kw in dt.get("tu_khoa", []):
            if bo_dau(kw) and bo_dau(kw) in t:
                return dt
    return None


def ensure_dir(p: Path) -> Path:
    p.mkdir(parents=True, exist_ok=True)
    return p


def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default)
