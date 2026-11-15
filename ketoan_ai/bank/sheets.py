"""Đọc tab sao kê trên Google Sheet bằng service account (gspread)."""
from __future__ import annotations

from pathlib import Path

from .. import utils


SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly", "https://www.googleapis.com/auth/drive.readonly"]


def _creds():
    import json

    from google.oauth2.service_account import Credentials

    info = utils.env("GOOGLE_SERVICE_ACCOUNT_INFO")
    if info:
        return Credentials.from_service_account_info(json.loads(info), scopes=SCOPES)
    sa_path = utils.env("GOOGLE_SERVICE_ACCOUNT_JSON", "config/service_account.json")
    sa_file = Path(sa_path) if Path(sa_path).is_absolute() else utils.ROOT / sa_path
    if not sa_file.exists():
        raise FileNotFoundError(f"Không thấy file service account {sa_file} và chưa đặt GOOGLE_SERVICE_ACCOUNT_INFO.")
    return Credentials.from_service_account_file(str(sa_file), scopes=SCOPES)


def doc_tab(sheet_id: str, tab: str) -> list[list]:
    import gspread

    gc = gspread.authorize(_creds())
    ws = gc.open_by_key(sheet_id).worksheet(tab)
    return ws.get_all_values()


def danh_sach_tab(sheet_id: str) -> list[str]:
    import gspread

    return [ws.title for ws in gspread.authorize(_creds()).open_by_key(sheet_id).worksheets()]
