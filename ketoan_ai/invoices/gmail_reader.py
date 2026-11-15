"""Đọc email hóa đơn từ hộp thư chung qua IMAP (Gmail: dùng Mật khẩu ứng dụng).

Trả về EmailHoaDon gồm: tiêu đề, người gửi, ghi chú đầu thư (tên công trình do bộ phận ghi),
file đính kèm (pdf/xml/zip/ảnh) và các link tra cứu hóa đơn tìm thấy trong nội dung.
"""
from __future__ import annotations

import email
import imaplib
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from email.header import decode_header, make_header
from email.message import Message

from bs4 import BeautifulSoup

from .. import utils


@dataclass
class EmailHoaDon:
    uid: str
    subject: str
    nguoi_gui: str
    ngay: str
    ghi_chu: str                       # dòng ghi chú đầu thư (tên công trình)
    body_text: str
    attachments: list[tuple[str, bytes]] = field(default_factory=list)
    links: list[str] = field(default_factory=list)


FWD_MARKERS = ["---------- Forwarded message", "Begin forwarded message", "-----Original Message-----",
               "Kính gửi", "Kinh gui", "From:", "Từ:", "Tu:", "Dear"]


def _decode(s) -> str:
    try:
        return str(make_header(decode_header(s or "")))
    except Exception:
        return s or ""


def _lay_ghi_chu(text: str) -> str:
    """Ghi chú = các dòng đầu thư do bộ phận gõ thêm, trước phần forward/nội dung gốc. Thường là tên công trình."""
    lines = []
    for ln in text.splitlines():
        s = ln.strip()
        if not s:
            if lines:
                break
            continue
        if any(s.startswith(m) or m in s for m in FWD_MARKERS):
            break
        lines.append(s)
        if len(lines) >= 3:
            break
    return " | ".join(lines)[:200]


def _tach_body(msg: Message) -> tuple[str, str]:
    text, html = "", ""
    for part in msg.walk():
        ct = part.get_content_type()
        if part.get_content_disposition() == "attachment":
            continue
        if ct == "text/plain" and not text:
            text = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "ignore")
        elif ct == "text/html" and not html:
            html = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "ignore")
    return text, html


def _tim_link(html: str, text: str) -> list[str]:
    cfg = utils.settings()["hoa_don"]
    patterns = [p.lower() for p in cfg.get("mau_link_tra_cuu", [])]
    urls: list[str] = []
    if html:
        soup = BeautifulSoup(html, "lxml")
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            label = a.get_text(" ", strip=True).lower()
            if href.startswith("http") and (any(p in href.lower() for p in patterns) or label in ("link", "tại đây", "tai day", "xem hóa đơn", "tra cứu")):
                urls.append(href)
    for m in re.findall(r"https?://[^\s<>\"']+", text or ""):
        if any(p in m.lower() for p in patterns):
            urls.append(m)
    seen, out = set(), []
    for u in urls:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def parse_message(uid: str, raw: bytes) -> EmailHoaDon:
    msg = email.message_from_bytes(raw)
    text, html = _tach_body(msg)
    plain = text or (BeautifulSoup(html, "lxml").get_text("\n") if html else "")
    atts = []
    for part in msg.walk():
        if part.get_content_disposition() == "attachment" or (part.get_filename() and part.get_content_maintype() != "text"):
            fn = _decode(part.get_filename())
            if fn and fn.lower().endswith((".pdf", ".xml", ".zip", ".png", ".jpg", ".jpeg")):
                atts.append((fn, part.get_payload(decode=True) or b""))
    return EmailHoaDon(uid=uid, subject=_decode(msg.get("Subject")), nguoi_gui=_decode(msg.get("From")),
                       ngay=msg.get("Date", ""), ghi_chu=_lay_ghi_chu(plain), body_text=plain[:5000],
                       attachments=atts, links=_tim_link(html, plain))


class GmailIMAP:
    def __init__(self):
        self.host = utils.env("KT_EMAIL_IMAP_HOST", "imap.gmail.com")
        self.user = utils.env("KT_EMAIL_USER")
        self.pwd = utils.env("KT_EMAIL_APP_PASSWORD").replace(" ", "")
        if not self.user or not self.pwd:
            raise RuntimeError("Chưa cấu hình KT_EMAIL_USER / KT_EMAIL_APP_PASSWORD trong .env")
        self.conn = imaplib.IMAP4_SSL(self.host)
        self.conn.login(self.user, self.pwd)

    def lay_email(self, so_ngay: int = 7, chi_chua_doc: bool = False, thu_muc: str = "INBOX") -> list[EmailHoaDon]:
        self.conn.select(f'"{thu_muc}"')
        since = (date.today() - timedelta(days=so_ngay)).strftime("%d-%b-%Y")
        crit = f'(SINCE {since}{" UNSEEN" if chi_chua_doc else ""})'
        _, data = self.conn.uid("search", None, crit)
        out = []
        for uid in data[0].split():
            _, msg_data = self.conn.uid("fetch", uid, "(BODY.PEEK[])")
            raw = msg_data[0][1]
            e = parse_message(uid.decode(), raw)
            # chỉ giữ email có dấu hiệu là hóa đơn
            if e.attachments or e.links or utils.find_keyword(e.subject, ["hoa don", "hóa đơn", "invoice", "e-invoice"]):
                out.append(e)
        return out

    def gan_nhan(self, uid: str, nhan: str):
        try:
            self.conn.create(nhan)
        except Exception:
            pass
        self.conn.uid("copy", uid, nhan)
        self.conn.uid("store", uid, "+FLAGS", "(\\Seen)")

    def dong(self):
        try:
            self.conn.logout()
        except Exception:
            pass
