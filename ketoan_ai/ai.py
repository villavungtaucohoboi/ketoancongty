"""Lớp gọi mô hình AI: đọc hóa đơn (PDF/ảnh), phân loại giao dịch ngân hàng, phân loại khoản mục chi phí.

Hỗ trợ 2 nhà cung cấp, chọn bằng biến môi trường KT_PROVIDER = "claude" | "openai":
  - claude : ANTHROPIC_API_KEY, model KT_MODEL (mặc định claude-opus-5-5)
  - openai : OPENAI_API_KEY,   model KT_OPENAI_MODEL (bỏ trống = tự chọn model mới nhất có trong tài khoản)

Mọi hàm đều có đường lùi (fallback) theo quy tắc từ khoá khi không có API key,
để app vẫn chạy được (ở mức kém thông minh hơn) và test được offline.
"""
from __future__ import annotations

import base64
import os
from typing import Literal

from pydantic import BaseModel, Field

from . import utils

DEFAULT_CLAUDE_MODEL = "claude-opus-5-5"


GROK_BASE_URL = "https://api.x.ai/v1"
DEFAULT_GROK_MODEL = "grok-4"


def provider() -> str:
    p = os.environ.get("KT_PROVIDER", "").lower()
    if p in ("claude", "anthropic") and os.environ.get("ANTHROPIC_API_KEY"):
        return "claude"
    if p == "openai" and os.environ.get("OPENAI_API_KEY"):
        return "openai"
    if p in ("grok", "xai") and os.environ.get("XAI_API_KEY"):
        return "grok"
    if os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"):
        return "claude"
    if os.environ.get("OPENAI_API_KEY"):
        return "openai"
    if os.environ.get("XAI_API_KEY"):
        return "grok"
    return "none"


def has_api_key() -> bool:
    return provider() != "none"


def model_name() -> str:
    if provider() == "claude":
        return os.environ.get("KT_MODEL", DEFAULT_CLAUDE_MODEL)
    if provider() == "openai":
        return os.environ.get("KT_OPENAI_MODEL") or _openai_auto_model()
    if provider() == "grok":
        return os.environ.get("KT_GROK_MODEL") or _grok_auto_model()
    return "(chưa cấu hình)"


def _grok_client():
    from openai import OpenAI  # xAI dùng giao thức tương thích OpenAI

    return OpenAI(api_key=os.environ["XAI_API_KEY"], base_url=os.environ.get("XAI_BASE_URL", GROK_BASE_URL))


_grok_models_cache: list[str] = []


def danh_sach_model_grok() -> list[str]:
    ids = [m.id for m in _grok_client().models.list()]
    ids = [i for i in ids if i.startswith("grok") and not any(x in i for x in ("image", "video", "audio", "tts", "code-fast"))]
    _grok_models_cache[:] = sorted(ids, reverse=True)
    return _grok_models_cache


def _grok_auto_model() -> str:
    try:
        ids = _grok_models_cache or danh_sach_model_grok()
    except Exception:  # noqa: BLE001
        return DEFAULT_GROK_MODEL
    for uu_tien in ("grok-5", "grok-4.2", "grok-4.1", "grok-4"):
        hit = [i for i in ids if i.startswith(uu_tien) and "mini" not in i and "fast" not in i]
        if hit:
            return sorted(hit)[-1]
    return ids[0] if ids else DEFAULT_GROK_MODEL


_openai_models_cache: list[str] = []


def danh_sach_model_openai() -> list[str]:
    """Liệt kê các model GPT có trong tài khoản OpenAI (để chọn trên giao diện)."""
    from openai import OpenAI

    ids = [m.id for m in OpenAI().models.list()]
    ids = [i for i in ids if i.startswith(("gpt", "o")) and not any(x in i for x in ("audio", "realtime", "image", "tts", "transcribe", "embedding", "search", "moderation", "codex", "whisper"))]
    _openai_models_cache[:] = sorted(ids, reverse=True)
    return _openai_models_cache


def _openai_auto_model() -> str:
    ids = _openai_models_cache or danh_sach_model_openai()
    for uu_tien in ("terra", "sol", "luna", "gpt-6", "gpt-5.5", "gpt-5", "gpt-4.1"):
        hit = [i for i in ids if uu_tien in i and "mini" not in i and "nano" not in i]
        if hit:
            return sorted(hit)[-1]
    return ids[0] if ids else "gpt-5"


# ---------------------------------------------------------------------------
# Schema dữ liệu trả về
# ---------------------------------------------------------------------------
class DongHang(BaseModel):
    stt: int = 0
    ten_hang: str
    dvt: str = ""
    so_luong: float = 0
    don_gia: float = 0
    thanh_tien: float = 0
    thue_suat: float | None = Field(default=None, description="Phần trăm, ví dụ 8 hoặc 10. null nếu không chịu thuế")
    tien_thue: float = 0


class HoaDon(BaseModel):
    mau_so: str = ""
    ky_hieu: str = ""
    so_hoa_don: str = ""
    ngay_hoa_don: str = Field(default="", description="dd/mm/yyyy")
    ten_nguoi_ban: str = ""
    mst_nguoi_ban: str = ""
    dia_chi_nguoi_ban: str = ""
    ten_nguoi_mua: str = ""
    mst_nguoi_mua: str = ""
    ma_cqt: str = ""
    tien_hang: float = 0
    tien_thue: float = 0
    tong_thanh_toan: float = 0
    hinh_thuc_thanh_toan: str = ""
    dong_hang: list[DongHang] = []
    nguon: str = ""           # xml / pdf / anh / link
    do_tin_cay: float = Field(default=1.0, description="0..1")
    ghi_chu: str = ""


class PhanLoaiHoaDon(BaseModel):
    khoan_muc_ma: str = Field(description="Mã khoản mục chi phí trong danh sách đã cho")
    cong_trinh_ma: str = Field(default="", description="Mã công trình trong danh sách, hoặc rỗng nếu không xác định")
    ly_do: str = ""
    do_tin_cay: float = Field(default=0.8, description="0..1")


class GiaoDichNganHang(BaseModel):
    stt: int
    loai: Literal["thu", "chi"]
    ma_doi_tuong: str = ""
    ten_doi_tuong: str = ""
    tk_doi_ung: str = Field(description="Tài khoản đối ứng theo TT200/TT133, ví dụ 331, 131, 6428, 3331, 334")
    cong_trinh_ma: str = ""
    khoan_muc_ma: str = ""
    dien_giai: str = Field(description="Diễn giải ngắn gọn, tiếng Việt có dấu, cho sổ kế toán")
    do_tin_cay: float = Field(default=0.8, description="0..1")


class DanhSachGiaoDich(BaseModel):
    giao_dich: list[GiaoDichNganHang]


# ---------------------------------------------------------------------------
# Gọi mô hình (trung lập nhà cung cấp)
# parts: [{"type":"text","text":...} | {"type":"pdf","data":bytes} | {"type":"image","data":bytes,"media_type":"image/png"}]
# ---------------------------------------------------------------------------
def _goi_mo_hinh(system: str, parts: list[dict], schema: type[BaseModel], max_tokens: int = 16000) -> BaseModel:
    p = provider()
    if p == "claude":
        return _goi_claude(system, parts, schema, max_tokens)
    if p == "openai":
        return _goi_openai(system, parts, schema, max_tokens)
    if p == "grok":
        return _goi_grok(system, parts, schema, max_tokens)
    raise RuntimeError("Chưa cấu hình API key (ANTHROPIC_API_KEY, OPENAI_API_KEY hoặc XAI_API_KEY).")


def _pdf_sang_anh(data: bytes, dpi: int = 110, toi_da: int = 4) -> list[bytes]:
    """Grok không nhận file PDF trực tiếp -> chuyển từng trang thành PNG bằng pdftoppm (Poppler)."""
    import subprocess
    import tempfile
    from pathlib import Path

    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / "in.pdf"
        src.write_bytes(data)
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-l", str(toi_da), str(src), str(Path(td) / "p")], check=True, capture_output=True)
        return [f.read_bytes() for f in sorted(Path(td).glob("p*.png"))]


def _goi_grok(system, parts, schema, max_tokens):
    content = []
    for pt in parts:
        if pt["type"] == "text":
            content.append({"type": "text", "text": pt["text"]})
        elif pt["type"] == "pdf":
            for png in _pdf_sang_anh(pt["data"]):
                content.append({"type": "image_url", "image_url": {"url": "data:image/png;base64," + base64.standard_b64encode(png).decode(), "detail": "high"}})
        elif pt["type"] == "image":
            content.append({"type": "image_url", "image_url": {"url": f"data:{pt['media_type']};base64," + base64.standard_b64encode(pt["data"]).decode(), "detail": "high"}})
    resp = _grok_client().chat.completions.parse(
        model=model_name(), messages=[{"role": "system", "content": system}, {"role": "user", "content": content}],
        response_format=schema, max_tokens=max_tokens,
    )
    msg = resp.choices[0].message
    if msg.parsed is None:
        raise RuntimeError(f"Grok không trả về dữ liệu hợp lệ: {(msg.content or '')[:200]}")
    return msg.parsed


def _goi_claude(system, parts, schema, max_tokens):
    import anthropic

    content = []
    for pt in parts:
        if pt["type"] == "text":
            content.append({"type": "text", "text": pt["text"]})
        elif pt["type"] == "pdf":
            content.append({"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                                           "data": base64.standard_b64encode(pt["data"]).decode()}})
        elif pt["type"] == "image":
            content.append({"type": "image", "source": {"type": "base64", "media_type": pt["media_type"],
                                                        "data": base64.standard_b64encode(pt["data"]).decode()}})
    resp = anthropic.Anthropic().messages.parse(
        model=os.environ.get("KT_MODEL", DEFAULT_CLAUDE_MODEL), max_tokens=max_tokens, system=system,
        messages=[{"role": "user", "content": content}], output_format=schema,
    )
    if resp.stop_reason == "refusal":
        raise RuntimeError("Claude từ chối xử lý tài liệu này (stop_reason=refusal).")
    return resp.parsed_output


def _goi_openai(system, parts, schema, max_tokens):
    from openai import OpenAI

    content = []
    for pt in parts:
        if pt["type"] == "text":
            content.append({"type": "input_text", "text": pt["text"]})
        elif pt["type"] == "pdf":
            content.append({"type": "input_file", "filename": "hoa_don.pdf",
                            "file_data": "data:application/pdf;base64," + base64.standard_b64encode(pt["data"]).decode()})
        elif pt["type"] == "image":
            content.append({"type": "input_image", "detail": "high",
                            "image_url": f"data:{pt['media_type']};base64," + base64.standard_b64encode(pt["data"]).decode()})
    resp = OpenAI().responses.parse(
        model=model_name(), instructions=system, input=[{"role": "user", "content": content}],
        text_format=schema, max_output_tokens=max_tokens,
    )
    if resp.output_parsed is None:
        raise RuntimeError(f"Mô hình không trả về dữ liệu hợp lệ: {resp.output_text[:200]}")
    return resp.output_parsed


# ---------------------------------------------------------------------------
# Prompt dùng chung
# ---------------------------------------------------------------------------
def _danh_muc_text() -> str:
    km = "\n".join(f"- {k['ma']}: {k['ten']} (từ khoá: {', '.join(k.get('tu_khoa', [])[:8])})" for k in utils.khoan_muc_list())
    ct = "\n".join(f"- {c['ma']}: {c['ten']} (còn gọi: {', '.join(c.get('bi_danh', []))})" for c in utils.cong_trinh_list())
    dt = "\n".join(f"- {d['ma']}: {d['ten']} (MST {d.get('mst') or '?'}; từ khoá: {', '.join(d.get('tu_khoa', []))})" for d in utils.doi_tuong_list())
    return f"DANH MỤC KHOẢN MỤC CHI PHÍ:\n{km}\n\nDANH MỤC CÔNG TRÌNH:\n{ct}\n\nDANH MỤC ĐỐI TƯỢNG (NCC/KH):\n{dt}"


SYSTEM_KE_TOAN = (
    "Bạn là trợ lý kế toán cho một công ty xây dựng tại Việt Nam, sử dụng phần mềm MISA SME.NET. "
    "Bạn đọc chứng từ cẩn thận, không bịa số liệu; nếu không chắc một trường nào thì để rỗng và hạ do_tin_cay (0..1). "
    "Số tiền trả về là số thuần (không dấu chấm phân cách nghìn). Ngày tháng dạng dd/mm/yyyy."
)


def _clamp(x: float) -> float:
    return max(0.0, min(1.0, float(x)))


# ---------------------------------------------------------------------------
# 1. Đọc hóa đơn từ PDF / ảnh
# ---------------------------------------------------------------------------
def doc_hoa_don_tu_file(data: bytes, media_type: str) -> HoaDon:
    """media_type: application/pdf | image/png | image/jpeg | image/webp"""
    if not has_api_key():
        raise RuntimeError("Chưa cấu hình API key nên không thể đọc hóa đơn PDF/ảnh bằng AI. "
                           "Hóa đơn dạng XML vẫn đọc được không cần AI.")
    part = {"type": "pdf", "data": data} if media_type == "application/pdf" else {"type": "image", "data": data, "media_type": media_type}
    hd: HoaDon = _goi_mo_hinh(SYSTEM_KE_TOAN, [part, {"type": "text", "text": (
        "Đây là một hóa đơn giá trị gia tăng điện tử của Việt Nam. Hãy trích xuất đầy đủ: "
        "mẫu số, ký hiệu, số hóa đơn, ngày, người bán (tên, MST, địa chỉ), người mua (tên, MST), "
        "mã của cơ quan thuế (nếu có), từng dòng hàng hóa dịch vụ (tên, ĐVT, số lượng, đơn giá, thành tiền, thuế suất), "
        "tiền hàng, tiền thuế, tổng thanh toán. Đặt nguon = 'pdf' hoặc 'anh'."
    )}], HoaDon)
    hd.nguon = hd.nguon or ("pdf" if media_type == "application/pdf" else "anh")
    hd.do_tin_cay = _clamp(hd.do_tin_cay)
    return hd


# ---------------------------------------------------------------------------
# 2. Phân loại khoản mục chi phí + công trình cho hóa đơn
# ---------------------------------------------------------------------------
def phan_loai_hoa_don(hd: HoaDon, ghi_chu_email: str = "") -> PhanLoaiHoaDon:
    ct = utils.match_cong_trinh(ghi_chu_email) if ghi_chu_email else None
    mo_ta = " ; ".join(d.ten_hang for d in hd.dong_hang) or hd.ghi_chu
    if not has_api_key():
        km = utils.match_khoan_muc(mo_ta + " " + hd.ten_nguoi_ban) or {"ma": ""}
        return PhanLoaiHoaDon(khoan_muc_ma=km["ma"], cong_trinh_ma=(ct or {}).get("ma", ""),
                              ly_do="Phân loại theo từ khoá (chưa có API key)", do_tin_cay=0.5 if km["ma"] else 0.2)
    out: PhanLoaiHoaDon = _goi_mo_hinh(SYSTEM_KE_TOAN + "\n\n" + _danh_muc_text(), [{"type": "text", "text": (
        f"Hóa đơn của người bán: {hd.ten_nguoi_ban} (MST {hd.mst_nguoi_ban}).\n"
        f"Hàng hóa/dịch vụ: {mo_ta}\n"
        f"Ghi chú của bộ phận gửi email (thường là tên công trình): {ghi_chu_email or '(không có)'}\n\n"
        "Chọn đúng 1 mã khoản mục chi phí và 1 mã công trình trong danh mục. "
        "Nếu ghi chú email nêu công trình thì ưu tiên theo ghi chú. Nếu không khớp công trình nào, để cong_trinh_ma rỗng."
    )}], PhanLoaiHoaDon, max_tokens=2000)
    out.do_tin_cay = _clamp(out.do_tin_cay)
    if ct and not out.cong_trinh_ma:
        out.cong_trinh_ma = ct["ma"]
    if out.khoan_muc_ma not in {k["ma"] for k in utils.khoan_muc_list()}:
        out.khoan_muc_ma = (utils.match_khoan_muc(mo_ta) or {"ma": ""})["ma"]
    return out


# ---------------------------------------------------------------------------
# 3. Phân loại giao dịch ngân hàng
# ---------------------------------------------------------------------------
def phan_loai_giao_dich(rows: list[dict]) -> list[GiaoDichNganHang]:
    """rows: [{stt, ngay, noi_dung, tien_ra, tien_vao, ma_cong_trinh_sheet}] -> danh sách đã phân loại.

    Thứ tự: (1) quy tắc từ khoá trong settings.yaml, (2) AI cho các dòng còn lại, (3) mặc định theo từ khoá danh mục.
    """
    cfg = utils.settings()["sao_ke"]
    rules = cfg.get("quy_tac", [])
    out: dict[int, GiaoDichNganHang] = {}
    pending = []
    for r in rows:
        loai = "chi" if utils.parse_money(r.get("tien_ra")) > 0 else "thu"
        nd = r.get("noi_dung", "")
        hit = next((q for q in rules if q.get("loai") == loai and utils.find_keyword(nd, q["tu_khoa"])), None)
        if hit:
            ct = utils.match_cong_trinh(nd)
            out[r["stt"]] = GiaoDichNganHang(
                stt=r["stt"], loai=loai, tk_doi_ung=str(hit["tk_doi_ung"]),
                khoan_muc_ma=(utils.match_khoan_muc(hit.get("khoan_muc", "")) or {}).get("ma", ""),
                cong_trinh_ma=(ct or {}).get("ma", ""), dien_giai=nd[:200], do_tin_cay=0.95)
        else:
            pending.append((r, loai))

    if pending and has_api_key():
        lines = []
        from .bank import rules as bank_rules

        for r, loai in pending:
            so_tien = utils.parse_money(r.get("tien_ra")) or utils.parse_money(r.get("tien_vao"))
            pt = bank_rules.phan_tich_noi_dung(r.get("noi_dung", ""), loai)
            goi_y = []
            if pt["cong_trinh"]:
                goi_y.append(f"CT gợi ý: {pt['cong_trinh']['ma']}")
            if pt["doi_tuong"]:
                goi_y.append(f"ĐT gợi ý: {pt['doi_tuong']['ma']}")
            lines.append(f"[{r['stt']}] {r.get('ngay','')} | {loai.upper()} | {utils.fmt_money(so_tien)} | {r.get('noi_dung','')}"
                         + (f" | Mã CT trên sheet: {r['ma_cong_trinh_sheet']}" if r.get("ma_cong_trinh_sheet") else "")
                         + (" | " + "; ".join(goi_y) if goi_y else ""))
        system = SYSTEM_KE_TOAN + "\n\n" + _danh_muc_text() + (
            "\n\nQUY ƯỚC TK ĐỐI ỨNG: tiền CHI trả nhà cung cấp -> 331; tiền THU từ khách hàng/chủ đầu tư -> 131; "
            "tạm ứng nhân viên -> 141; lương -> 334; BHXH -> 3383; thuế -> 3331/3334; phí ngân hàng -> 6428; "
            "lãi tiền gửi -> 515; rút/nộp tiền mặt -> 1111; vay ngân hàng -> 341; trả lãi vay -> 635. "
            "Mã viết tắt hay gặp: TT = thanh toán, CT = công trình, HD/HĐ = hợp đồng, NCC = nhà cung cấp, "
            "VP = văn phòng, NVL = nguyên vật liệu, CTY = công ty, HCIC = công ty mình. Giữ nguyên 'loai' như đã cho.")
        ds: DanhSachGiaoDich = _goi_mo_hinh(system, [{"type": "text", "text": (
            "Phân loại từng giao dịch sao kê ngân hàng sau (mỗi dòng: [stt] ngày | loại | số tiền | nội dung). "
            "Với mỗi dòng: xác định tài khoản đối ứng, mã đối tượng (nếu khớp danh mục, nếu không thì để rỗng và điền ten_doi_tuong "
            "là tên NCC/KH đọc được từ nội dung), mã công trình (nếu nội dung hoặc Mã CT trên sheet nêu), khoản mục chi phí "
            "(nếu là chi phí/vật tư), và diễn giải tiếng Việt có dấu ngắn gọn. Trả về đúng và đủ các stt đã cho.\n\n" + "\n".join(lines)
        )}], DanhSachGiaoDich)
        loai_cua = {r["stt"]: loai for r, loai in pending}
        for g in ds.giao_dich:
            if g.stt in loai_cua:
                g.loai = loai_cua[g.stt]  # không cho AI đổi chiều thu/chi
                g.do_tin_cay = _clamp(g.do_tin_cay)
                out[g.stt] = g

    from .bank import rules as bank_rules

    for r, loai in pending:
        if r["stt"] not in out:
            nd = r.get("noi_dung", "")
            pt = bank_rules.phan_tich_noi_dung(nd, loai)
            ct = pt["cong_trinh"] or (utils.match_cong_trinh(r.get("ma_cong_trinh_sheet") or "") if r.get("ma_cong_trinh_sheet") else None)
            km = utils.match_khoan_muc(pt["muc_dich"] or nd) if loai == "chi" else None
            if loai == "chi" and ct and ct["ma"] == "VP" and not km:
                km = next((k for k in utils.khoan_muc_list() if "quản lý chung" in k["ten"]), None)
            out[r["stt"]] = GiaoDichNganHang(
                stt=r["stt"], loai=loai, tk_doi_ung=cfg["tk_doi_ung_mac_dinh"][loai],
                ma_doi_tuong=(pt["doi_tuong"] or {}).get("ma", ""), ten_doi_tuong=pt["ten_doi_tuong"],
                cong_trinh_ma=(ct or {}).get("ma", ""), khoan_muc_ma=(km or {}).get("ma", ""),
                dien_giai=pt["dien_giai"], do_tin_cay=pt["do_tin_cay"])
    return [out[r["stt"]] for r in rows]
