"""Mở link tra cứu hóa đơn bằng trình duyệt (Playwright) và lấy về PDF/XML hoặc ảnh chụp trang.

Mỗi nhà cung cấp HĐĐT (EasyInvoice, M-Invoice, VNPT, Viettel, BKAV...) có trang tra cứu khác nhau.
Chiến lược chung: mở trang -> nếu trình duyệt tải xuống file (pdf/xml) thì lấy file -> nếu có nút
"Tải PDF"/"Tải XML" thì bấm -> nếu không, chụp ảnh toàn trang để AI đọc.
"""
from __future__ import annotations

import re


def tai_tu_link(url: str, timeout_ms: int = 30000) -> tuple[bytes, str]:
    """Trả về (dữ liệu, media_type) với media_type ∈ application/pdf | text/xml | image/png."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        import os
        launch_kw = {"headless": True}
        if os.environ.get("KT_CHROMIUM_PATH"):  # dùng Chromium có sẵn trên máy thay vì 'playwright install'
            launch_kw["executable_path"] = os.environ["KT_CHROMIUM_PATH"]
        browser = p.chromium.launch(**launch_kw)
        ctx = browser.new_context(accept_downloads=True, viewport={"width": 1280, "height": 1800})
        page = ctx.new_page()
        downloaded: list[tuple[bytes, str]] = []

        def on_download(d):
            path = d.path()
            data = open(path, "rb").read()
            name = (d.suggested_filename or "").lower()
            mt = "application/pdf" if name.endswith(".pdf") or data[:4] == b"%PDF" else ("text/xml" if name.endswith(".xml") or data[:5] == b"<?xml" else "application/octet-stream")
            downloaded.append((data, mt))

        page.on("download", on_download)
        try:
            resp = page.goto(url, wait_until="networkidle", timeout=timeout_ms)
        except Exception:
            resp = None
        if resp is not None:
            ct = (resp.headers.get("content-type") or "").lower()
            if "application/pdf" in ct or "xml" in ct:
                # Chromium hiển thị PDF bằng trình xem riêng nên resp.body() có thể là HTML -> tải lại bằng request thuần
                try:
                    data = ctx.request.get(resp.url).body()
                except Exception:
                    data = resp.body()
                if data[:4] == b"%PDF":
                    browser.close()
                    return data, "application/pdf"
                if data[:5] == b"<?xml" or b"<HDon" in data[:2000]:
                    browser.close()
                    return data, "text/xml"
        if downloaded:
            browser.close()
            return downloaded[0]
        # thử bấm nút tải PDF / XML
        for pat in [r"tải.*pdf", r"download.*pdf", r"xem.*pdf", r"pdf", r"tải.*xml", r"xml"]:
            try:
                btn = page.get_by_role("button", name=re.compile(pat, re.I)).first
                if btn.count() == 0:
                    btn = page.get_by_role("link", name=re.compile(pat, re.I)).first
                if btn.count():
                    with page.expect_download(timeout=8000) as dl:
                        btn.click()
                    d = dl.value
                    data = open(d.path(), "rb").read()
                    browser.close()
                    return data, ("application/pdf" if data[:4] == b"%PDF" else "text/xml")
            except Exception:
                continue
        # iframe PDF?
        for fr in page.frames:
            if fr.url.lower().endswith(".pdf") or ".pdf" in fr.url.lower():
                try:
                    r = ctx.request.get(fr.url)
                    data = r.body()
                    if data[:4] == b"%PDF":
                        browser.close()
                        return data, "application/pdf"
                except Exception:
                    pass
        png = page.screenshot(full_page=True, type="png")
        browser.close()
        return png, "image/png"
