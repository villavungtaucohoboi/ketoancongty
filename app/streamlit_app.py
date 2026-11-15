"""Giao diện web cho kế toán: streamlit run app/streamlit_app.py"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ketoan_ai import ai, utils  # noqa: E402

st.set_page_config(page_title="Trợ lý kế toán AI", page_icon="🧾", layout="wide")

# ---- Nạp bí mật từ Streamlit Cloud (Settings > Secrets) vào biến môi trường
import hmac  # noqa: E402
import os  # noqa: E402

try:
    for _k, _v in st.secrets.items():
        if isinstance(_v, (str, int, float)) and not os.environ.get(_k):
            os.environ[_k] = str(_v)
except Exception:  # noqa: BLE001  (không có secrets.toml khi chạy máy nội bộ)
    pass


def _dang_nhap() -> bool:
    """Bảo vệ webapp bằng mật khẩu KT_APP_PASSWORD. Bỏ trống biến này = không yêu cầu đăng nhập (chỉ dùng trên máy nội bộ)."""
    mk = os.environ.get("KT_APP_PASSWORD", "")
    if not mk or st.session_state.get("da_dang_nhap"):
        return True
    st.title("🧾 Trợ lý kế toán AI")
    with st.form("login"):
        nhap = st.text_input("Mật khẩu truy cập", type="password")
        ok = st.form_submit_button("Đăng nhập", type="primary")
    if ok:
        if hmac.compare_digest(nhap.encode(), mk.encode()):
            st.session_state["da_dang_nhap"] = True
            st.rerun()
        st.error("Sai mật khẩu.")
    return False


if not _dang_nhap():
    st.stop()

st.title("🧾 Trợ lý kế toán AI")

SAMPLES = ROOT / "samples"

with st.sidebar:
    st.header("⚙️ Cấu hình AI")
    st.caption("Khóa nhập ở đây chỉ dùng trong phiên này. Khóa cố định đặt trong .env hoặc Secrets của máy chủ.")
    if st.session_state.get("da_dang_nhap") and st.button("Đăng xuất"):
        st.session_state.clear()
        st.rerun()
    _provs = ["Claude (Anthropic)", "Grok (xAI)", "OpenAI (ChatGPT)"]
    _cur = os.environ.get("KT_PROVIDER", "claude")
    prov = st.radio("Nhà cung cấp", _provs, index={"grok": 1, "openai": 2}.get(_cur, 0))
    if prov.startswith("Grok"):
        key = st.text_input("XAI_API_KEY", value=os.environ.get("XAI_API_KEY", ""), type="password", placeholder="xai-...",
                            help="Lấy tại console.x.ai → API Keys. Gói SuperGrok / grok.com không kèm API.")
        if key:
            os.environ["XAI_API_KEY"] = key.strip()
            os.environ["KT_PROVIDER"] = "grok"
        if st.button("Lấy danh sách model Grok"):
            try:
                st.session_state["grok_models"] = ai.danh_sach_model_grok()
            except Exception as ex:  # noqa: BLE001
                st.error(f"Không kết nối được xAI: {ex}")
        models = st.session_state.get("grok_models", [])
        if models:
            cur = os.environ.get("KT_GROK_MODEL") or ai.model_name()
            chon = st.selectbox("Model", models, index=models.index(cur) if cur in models else 0)
            os.environ["KT_GROK_MODEL"] = chon
    elif prov.startswith("OpenAI"):
        key = st.text_input("OPENAI_API_KEY", value=os.environ.get("OPENAI_API_KEY", ""), type="password", placeholder="sk-proj-...")
        if key:
            os.environ["OPENAI_API_KEY"] = key.strip()
            os.environ["KT_PROVIDER"] = "openai"
        if st.button("Lấy danh sách model"):
            try:
                st.session_state["oa_models"] = ai.danh_sach_model_openai()
            except Exception as ex:  # noqa: BLE001
                st.error(f"Không kết nối được OpenAI: {ex}")
        models = st.session_state.get("oa_models", [])
        if models:
            cur = os.environ.get("KT_OPENAI_MODEL") or ai.model_name()
            chon = st.selectbox("Model", models, index=models.index(cur) if cur in models else 0)
            os.environ["KT_OPENAI_MODEL"] = chon
    else:
        key = st.text_input("ANTHROPIC_API_KEY", value=os.environ.get("ANTHROPIC_API_KEY", ""), type="password", placeholder="sk-ant-...")
        if key:
            os.environ["ANTHROPIC_API_KEY"] = key.strip()
            os.environ["KT_PROVIDER"] = "claude"
        os.environ["KT_MODEL"] = st.text_input("Model Claude", os.environ.get("KT_MODEL", ai.DEFAULT_CLAUDE_MODEL))
    if ai.has_api_key():
        try:
            st.success(f"AI đang bật: {ai.provider()} / {ai.model_name()}")
        except Exception as ex:  # noqa: BLE001
            st.warning(f"Có key nhưng chưa chọn được model: {ex}")
    else:
        st.warning("Chưa có API key → app chạy chế độ DEMO theo từ khoá (không đọc được PDF/ảnh).")
    st.divider()
    st.markdown("**Dữ liệu mẫu** có sẵn ở mỗi tab, bấm nút *Dùng dữ liệu mẫu* để xem luồng xử lý.")

tab1, tab2, tab3, tab4 = st.tabs(["1. Sao kê ngân hàng", "2. Hóa đơn mua vào", "3. Đối chiếu bảng kê", "4. Tính lương"])


def _save_upload(up, sub="uploads") -> Path:
    d = utils.ensure_dir(utils.DATA_DIR / sub)
    p = d / f"{datetime.now():%Y%m%d_%H%M%S}_{up.name}"
    p.write_bytes(up.getbuffer())
    return p


def _download(path: Path, label: str):
    with open(path, "rb") as f:
        st.download_button(label, f, file_name=path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------- Tab 1 ----------------
with tab1:
    from ketoan_ai.bank import pipeline as bank_pipeline

    cfg = utils.settings()["sao_ke"]
    c1, c2, c3 = st.columns([2, 1, 1])
    tab_name = c1.selectbox("Ngân hàng (tab trên Google Sheet)", list(cfg["tabs"].keys()))
    tu = c2.date_input("Từ ngày", value=None, format="DD/MM/YYYY")
    den = c3.date_input("Đến ngày", value=None, format="DD/MM/YYYY")
    nguon = st.radio("Nguồn dữ liệu", ["Google Sheet (cấu hình trong .env)", "Tải file sao kê lên (xlsx/csv)"], horizontal=True)
    up = st.file_uploader("File sao kê", type=["xlsx", "xls", "csv"]) if nguon.startswith("Tải") else None
    b1, b2 = st.columns([1, 1])
    run_real = b1.button("▶ Đọc & phân loại sao kê", type="primary")
    run_demo = b2.button("🧪 Dùng sao kê mẫu (ACB, 12 giao dịch)")
    if run_real or run_demo:
        try:
            with st.spinner("Đang đọc và phân loại..."):
                src = SAMPLES / "sao_ke_ACB_mau.xlsx" if run_demo else (_save_upload(up) if up else None)
                df, out = bank_pipeline.chay("ACB_CTY" if run_demo else tab_name, src, None if run_demo else tu, None if run_demo else den)
            st.session_state["bank_df"], st.session_state["bank_out"] = df, out
        except Exception as ex:  # noqa: BLE001
            st.error(str(ex))
    if "bank_df" in st.session_state:
        df = st.session_state["bank_df"]
        st.caption("Dòng tô vàng = AI chưa chắc (tin cậy < 70%), cần kế toán xem lại trước khi nhập.")
        st.dataframe(df.style.apply(lambda r: ["background-color:#fff2cc" if r["Tin cậy"] < 0.7 else ""] * len(r), axis=1)
                     .format({"Số tiền": "{:,.0f}", "Tin cậy": "{:.0%}"}), use_container_width=True, height=420)
        _download(st.session_state["bank_out"], "⬇ Tải file nhập khẩu MISA (Thu/Chi tiền gửi)")
        st.info("Vào MISA SME.NET → Tệp → Nhập khẩu từ Excel → chọn 'Thu tiền gửi' / 'Chi tiền gửi' → chọn file này → ghép cột → Nhập khẩu.")

# ---------------- Tab 2 ----------------
with tab2:
    from ketoan_ai.invoices import pipeline as inv_pipeline, store

    st.subheader("Lấy hóa đơn")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Từ email chung** (cấu hình trong .env)")
        so_ngay = st.number_input("Số ngày gần nhất", 1, 60, 7)
        dung_link = st.checkbox("Mở link tra cứu khi email không có file đính kèm (chậm hơn)", value=True)
        if st.button("📧 Quét email hóa đơn"):
            try:
                with st.spinner("Đang quét email..."):
                    tk = inv_pipeline.chay_tu_email(int(so_ngay), dung_link=dung_link)
                st.success(f"Email: {tk['email']} | Hóa đơn mới: {tk['moi']} | Trùng (đã có): {tk['trung']}")
                for cb in tk["canh_bao"]:
                    st.warning(cb)
            except Exception as ex:  # noqa: BLE001
                st.error(str(ex))
    with c2:
        st.markdown("**Hoặc tải file hóa đơn lên** (XML / PDF / ZIP / ảnh)")
        ghi_chu = st.text_input("Tên công trình (nếu biết)", "")
        ups = st.file_uploader("Chọn file", type=["xml", "pdf", "zip", "png", "jpg", "jpeg"], accept_multiple_files=True)
        d1, d2 = st.columns(2)
        if d1.button("🧪 Nạp email mẫu (.eml có zip)"):
            from ketoan_ai.invoices import extractor, store as _store
            from ketoan_ai.invoices.gmail_reader import parse_message
            e = parse_message("demo-eml", (SAMPLES / "email_hoa_don_mau.eml").read_bytes())
            kq, cb = extractor.trich_xuat_tu_email(e, dung_link=False)
            n = sum(1 for hd, pl in kq if _store.luu(hd.model_dump(), pl.model_dump(), e.uid, e.subject))
            st.success(f"Email: '{e.subject[:60]}...' | ghi chú đầu thư: '{e.ghi_chu}' | hóa đơn mới: {n}")
            for c in cb:
                st.warning(c)
        if d2.button("🧪 Nạp 2 hóa đơn XML mẫu"):
            tk = inv_pipeline.chay_tu_thu_muc(SAMPLES, "Nhà B An Khánh")
            st.success(f"File: {tk['file']} | Hóa đơn mới: {tk['moi']} | Trùng: {tk['trung']}")
        if st.button("📂 Đọc các file đã tải") and ups:
            d = utils.ensure_dir(utils.DATA_DIR / "inbox_hoadon" / f"{datetime.now():%Y%m%d_%H%M%S}")
            for u in ups:
                (d / u.name).write_bytes(u.getbuffer())
            try:
                with st.spinner("Đang đọc hóa đơn..."):
                    tk = inv_pipeline.chay_tu_thu_muc(d, ghi_chu)
                st.success(f"File: {tk['file']} | Hóa đơn mới: {tk['moi']} | Trùng: {tk['trung']}")
                for cb in tk["canh_bao"]:
                    st.warning(cb)
            except Exception as ex:  # noqa: BLE001
                st.error(str(ex))

    st.subheader("Hàng đợi chờ duyệt")
    items = store.danh_sach("cho_duyet")
    if not items:
        st.info("Chưa có hóa đơn nào chờ duyệt.")
    else:
        km_opts = [""] + [k["ma"] for k in utils.khoan_muc_list()]
        ct_opts = [""] + [c["ma"] for c in utils.cong_trinh_list()]
        rows = []
        for it in items:
            hd, pl = it["hd"], it["pl"]
            rows.append({"Khóa": it["khoa"], "Ngày HĐ": hd["ngay_hoa_don"], "Ký hiệu": hd["ky_hieu"], "Số HĐ": hd["so_hoa_don"],
                         "Người bán": hd["ten_nguoi_ban"], "MST": hd["mst_nguoi_ban"],
                         "Hàng hóa": "; ".join(d["ten_hang"] for d in hd["dong_hang"])[:120],
                         "Tiền hàng": hd["tien_hang"], "Thuế": hd["tien_thue"], "Tổng": hd["tong_thanh_toan"],
                         "Khoản mục": pl["khoan_muc_ma"], "Công trình": pl["cong_trinh_ma"], "Nguồn": hd["nguon"],
                         "Tin cậy": min(hd.get("do_tin_cay", 1), pl.get("do_tin_cay", 1)), "Xuất": True})
        df = pd.DataFrame(rows)
        edited = st.data_editor(
            df, use_container_width=True, hide_index=True, height=380, disabled=[c for c in df.columns if c not in ("Khoản mục", "Công trình", "Xuất")],
            column_config={"Khoản mục": st.column_config.SelectboxColumn(options=km_opts), "Công trình": st.column_config.SelectboxColumn(options=ct_opts),
                           "Tiền hàng": st.column_config.NumberColumn(format="%d"), "Thuế": st.column_config.NumberColumn(format="%d"),
                           "Tổng": st.column_config.NumberColumn(format="%d"), "Tin cậy": st.column_config.ProgressColumn(min_value=0, max_value=1)})
        c1, c2 = st.columns(2)
        if c1.button("💾 Lưu sửa đổi khoản mục / công trình"):
            for _, r in edited.iterrows():
                store.cap_nhat(r["Khóa"], phan_loai={"khoan_muc_ma": r["Khoản mục"], "cong_trinh_ma": r["Công trình"] or ""})
            st.success("Đã lưu.")
            st.rerun()
        if c2.button("⬇ Duyệt & xuất file nhập MISA (Mua hàng)", type="primary"):
            for _, r in edited.iterrows():
                store.cap_nhat(r["Khóa"], phan_loai={"khoan_muc_ma": r["Khoản mục"], "cong_trinh_ma": r["Công trình"] or ""},
                               trang_thai="cho_duyet" if r["Xuất"] else "bo_qua")
            out, n = inv_pipeline.xuat_misa()
            if out:
                st.success(f"Đã xuất {n} hóa đơn.")
                _download(out, "⬇ Tải file MISA Mua hàng")
                st.info("MISA SME.NET → Tệp → Nhập khẩu từ Excel → 'Mua hàng' (Hóa đơn mua hàng) → chọn file → ghép cột → Nhập khẩu. "
                        "Các NCC chưa có trong danh mục (tô vàng) cần thêm vào Danh mục > Đối tượng trước.")
    da_xuat = store.danh_sach("da_xuat")
    if da_xuat:
        with st.expander(f"Đã xuất sang MISA ({len(da_xuat)})"):
            st.dataframe(pd.DataFrame([{"Ngày": i["hd"]["ngay_hoa_don"], "Số HĐ": i["hd"]["so_hoa_don"], "Người bán": i["hd"]["ten_nguoi_ban"],
                                        "Tổng": i["hd"]["tong_thanh_toan"], "Khoản mục": i["pl"]["khoan_muc_ma"], "Công trình": i["pl"]["cong_trinh_ma"]} for i in da_xuat]),
                         use_container_width=True)

# ---------------- Tab 3 ----------------
with tab3:
    from ketoan_ai import misa_db
    from ketoan_ai.reconcile import compare

    nguon_misa = st.radio("Dữ liệu MISA lấy từ", ["Đọc thẳng CSDL MISA (máy trong mạng công ty)", "Tải file bảng kê xuất từ MISA"],
                          index=0 if misa_db.da_cau_hinh() else 1, horizontal=True)
    doc_thang = nguon_misa.startswith("Đọc thẳng")
    c1, c2 = st.columns(2)
    if doc_thang:
        if not misa_db.da_cau_hinh():
            c1.warning("Chưa cấu hình kết nối MISA (KT_MISA_DB trong .env). Chức năng này chỉ chạy khi app chạy trên máy "
                       "trong mạng công ty, không chạy trên Streamlit Cloud. Xem mục 'Kết nối MISA' trong hướng dẫn.")
        _d = datetime.now().date().replace(day=1)
        _thang_truoc = (_d - pd.Timedelta(days=1))
        tu_dc = c1.date_input("Từ ngày (ngày hóa đơn)", value=_thang_truoc.replace(day=1), format="DD/MM/YYYY", key="dc_tu")
        den_dc = c1.date_input("Đến ngày", value=_thang_truoc, format="DD/MM/YYYY", key="dc_den")
        c1.caption("Chỉ đọc (SELECT), dữ liệu MISA không rời khỏi máy nội bộ.")
        f_misa = None
    else:
        c1.markdown("**Bảng kê MISA**: Báo cáo → Thuế → Bảng kê HĐ mua vào → Xuất Excel.")
        f_misa = c1.file_uploader("Bảng kê MISA", type=["xlsx", "xls", "csv"], key="misa")
    c2.markdown("**Bảng kê Thuế**: hoadondientu.gdt.gov.vn → Tra cứu → Hóa đơn mua vào → Xuất Excel.")
    f_thue = c2.file_uploader("Bảng kê Thuế", type=["xlsx", "xls", "csv"], key="thue")
    e1, e2 = st.columns(2)
    run_real = e1.button("▶ Đối chiếu", type="primary")
    run_demo = e2.button("🧪 Dùng bảng kê mẫu")
    if run_real and not f_thue:
        st.warning("Hãy tải bảng kê Thuế lên trước.")
    if run_demo or (run_real and f_thue and (doc_thang or f_misa)):
        try:
            if run_demo:
                kq = compare.doi_chieu(SAMPLES / "bang_ke_misa_mau.xlsx", SAMPLES / "bang_ke_thue_mau.xlsx")
            elif doc_thang:
                with st.spinner("Đang đọc hóa đơn mua vào từ CSDL MISA..."):
                    df_misa = misa_db.bang_ke_mua_vao(tu_dc, den_dc)
                st.info(f"Đọc được {len(df_misa)} dòng hóa đơn mua vào từ MISA ({tu_dc:%d/%m/%Y} – {den_dc:%d/%m/%Y}).")
                kq = compare.doi_chieu(df_misa, _save_upload(f_thue))
            else:
                kq = compare.doi_chieu(_save_upload(f_misa), _save_upload(f_thue))
            st.session_state["dc"] = kq
        except Exception as ex:  # noqa: BLE001
            st.error(str(ex))
    if "dc" in st.session_state:
        kq = st.session_state["dc"]
        st.dataframe(kq["Tong hop"], use_container_width=True, hide_index=True)
        for name in ("Chi co tren Thue", "Chi co tren MISA", "Lech tien"):
            with st.expander(f"{name} ({len(kq[name])})", expanded=len(kq[name]) > 0):
                st.dataframe(kq[name], use_container_width=True, hide_index=True)
        out = compare.xuat_bao_cao(kq, utils.DATA_DIR / "bao_cao" / f"DoiChieu_{datetime.now():%Y%m%d_%H%M}.xlsx")
        _download(out, "⬇ Tải báo cáo đối chiếu")

# ---------------- Tab 4 ----------------
with tab4:
    from ketoan_ai.payroll import engine, mau_cong_ty

    che_do = st.radio("Cách tính", ["Theo file bảng lương của công ty (khuyên dùng)", "Bảng lương mẫu chung của app"], horizontal=True)
    if che_do.startswith("Theo file"):
        st.markdown(
            "1. Mở file bảng lương **tháng trước** bằng Excel, nhân sự cập nhật sheet **Chấm công + đánh giá hiệu quả** cho tháng mới "
            "(ngày công, điểm/hệ số hiệu quả, nghỉ phép, làm thêm), lưu lại.\n"
            "2. Tải file đó lên dưới đây, chọn tháng lương, bấm **Tạo bảng lương**.\n"
            "3. App chép số liệu sang sheet **Bảng lương**, đổi tiêu đề tháng, cập nhật bảng tổng hợp. "
            "**Không đổi bất kỳ công thức, hệ số nào**; Excel tự tính lại khi mở file.")
        k1, k2 = st.columns([3, 1])
        f_mau = k1.file_uploader("File bảng lương (đã cập nhật chấm công tháng mới)", type=["xlsx"], key="mau_cty")
        _t = datetime.now()
        thang_mau = k2.text_input("Tháng lương", f"{_t.month:02d}/{_t.year}", key="thang_mau")
        if st.button("▶ Tạo bảng lương", type="primary", key="btn_mau") and f_mau:
            try:
                src = _save_upload(f_mau)
                out = utils.DATA_DIR / "bao_cao" / f"Bang_luong_thang_{thang_mau.replace('/', '.')}.xlsx"
                with st.spinner("Đang chép số liệu và tính thử (khoảng 10–30 giây)..."):
                    kq = mau_cong_ty.cap_nhat(src, thang_mau, out)
                st.session_state["luong_mau"] = (kq, out)
            except Exception as ex:  # noqa: BLE001
                st.error(str(ex))
        if "luong_mau" in st.session_state:
            kq, out = st.session_state["luong_mau"]
            if kq.thang_cu and kq.thang_cu != kq.thang_moi:
                st.success(f"Đã tạo bảng lương tháng {kq.thang_moi} từ file tháng {kq.thang_cu}.")
            for cb in kq.canh_bao:
                st.warning(cb)
            st.markdown(f"**Ô đã cập nhật từ chấm công: {len(kq.thay_doi)}** (các ô khác giữ nguyên)")
            if len(kq.thay_doi):
                st.dataframe(kq.thay_doi, use_container_width=True, hide_index=True)
            if kq.da_tinh:
                cols = st.columns(len(kq.tong) + 1)
                for col, (k, v) in zip(cols, kq.tong.items()):
                    col.metric(f"Thực nhận {k}", utils.fmt_money(v))
                cols[-1].metric("Tổng chi lương", utils.fmt_money(sum(kq.tong.values())))
                with st.expander("Thực nhận từng người"):
                    st.dataframe(kq.bang_luong, use_container_width=True, hide_index=True)
            else:
                st.info("Máy chủ không có LibreOffice nên chưa xem trước được số tiền. Tải file về và mở bằng Excel, Excel tự tính đúng theo công thức.")
            _download(out, f"⬇ Tải bảng lương tháng {kq.thang_moi}")
        st.caption("Muốn dùng bảng lương mẫu chung (không theo file công ty), chọn cách tính thứ hai ở trên.")
    else:
        st.markdown("Quy tắc tính lương nằm trong `config/luong_rules.yaml` (ngày công chuẩn, hệ số tăng ca, % thưởng theo xếp loại, "
                    "bảo hiểm, giảm trừ, biểu thuế). Mẫu file đầu vào: xem thư mục `samples/`.")
        c1, c2, c3 = st.columns(3)
        f_ns = c1.file_uploader("Bảng nhân sự (lương CB, phụ cấp, NPT)", type=["xlsx", "csv"], key="ns")
        f_cc = c2.file_uploader("Bảng chấm công", type=["xlsx", "csv"], key="cc")
        f_dg = c3.file_uploader("Bảng đánh giá (xếp loại)", type=["xlsx", "csv"], key="dg")
        thang = st.text_input("Tháng lương", f"{datetime.now():%m/%Y}")
        g1, g2 = st.columns(2)
        run_real = g1.button("▶ Tính lương", type="primary")
        run_demo = g2.button("🧪 Dùng file mẫu (3 nhân viên)")
        if run_demo or (run_real and f_ns and f_cc):
            try:
                df = engine.tinh_luong(SAMPLES / "nhan_su_mau.xlsx", SAMPLES / "cham_cong_mau.xlsx", SAMPLES / "danh_gia_mau.xlsx") if run_demo \
                    else engine.tinh_luong(_save_upload(f_ns), _save_upload(f_cc), _save_upload(f_dg) if f_dg else None)
                st.session_state["luong"] = df
            except Exception as ex:  # noqa: BLE001
                st.error(str(ex))
        if "luong" in st.session_state:
            df = st.session_state["luong"]
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.metric("Tổng thực lĩnh", utils.fmt_money(df["THỰC LĨNH"].sum()))
            out = engine.xuat_bang_luong(df, utils.DATA_DIR / "bao_cao" / f"BangLuong_{thang.replace('/', '_')}.xlsx", thang)
            _download(out, "⬇ Tải bảng lương")
