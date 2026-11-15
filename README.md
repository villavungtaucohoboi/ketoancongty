# Trợ lý kế toán AI (ketoan-ai)

App nội bộ giúp kế toán tự động hoá 4 việc lặp lại hàng ngày/tháng, đầu ra là **file Excel nhập khẩu thẳng vào MISA SME.NET** (và báo cáo Excel):

| # | Phân hệ | Đầu vào | AI làm gì | Đầu ra |
|---|---------|---------|-----------|--------|
| 1 | Sao kê ngân hàng | Google Sheet "Sao kê ngân hàng 2026" (tab từng ngân hàng) hoặc file sao kê xlsx/csv | Đọc từng giao dịch, xác định thu/chi, TK đối ứng, đối tượng, công trình, khoản mục, diễn giải | File MISA *Thu tiền gửi / Chi tiền gửi* |
| 2 | Hóa đơn mua vào | Email chung nhận hóa đơn (file XML/PDF/ZIP đính kèm hoặc link tra cứu) | Đọc hóa đơn, lấy danh mục hàng hóa, **phân loại khoản mục chi phí**, lấy **tên công trình** từ ghi chú đầu email, chống trùng | Hàng đợi duyệt + file MISA *Mua hàng* |
| 3 | Đối chiếu bảng kê | Bảng kê mua vào xuất từ MISA + bảng kê tải từ hoadondientu.gdt.gov.vn | So khớp theo MST + số HĐ, tìm thiếu/thừa/lệch tiền | Báo cáo đối chiếu Excel |
| 4 | Tính lương | Bảng nhân sự + bảng chấm công + bảng đánh giá (Excel) | Tính lương theo quy tắc cấu hình (ngày công, tăng ca, thưởng KPI, BH, thuế TNCN) | Bảng lương Excel |

Hướng dẫn sử dụng, đấu nối API và triển khai webapp: tài liệu "Hướng dẫn sử dụng Trợ lý kế toán AI" (https://claude.ai/artifact/2B4gSa6Ed7ukoiq8fxavbr) và **[HUONG_DAN_SETUP.md](HUONG_DAN_SETUP.md)**.

## Đưa lên web

- Streamlit Cloud: đẩy thư mục này lên GitHub (private), tạo app trỏ tới `app/streamlit_app.py`, dán Secrets theo `.streamlit/secrets.toml.example`.
- Docker: `cp .env.example .env` (điền khóa, đặt `KT_APP_PASSWORD`) rồi `docker compose up -d --build`, mở cổng 8501 qua HTTPS.

## Chạy nhanh

```bash
pip install -r requirements.txt
cp .env.example .env          # rồi điền API key, email, sheet id
python samples/tao_du_lieu_mau.py
streamlit run app/streamlit_app.py
```

Hoặc dùng dòng lệnh:

```bash
python -m ketoan_ai.cli sao-ke --tab ACB_CTY --file samples/sao_ke_ACB_mau.xlsx
python -m ketoan_ai.cli hoa-don thu-muc samples --ghi-chu "Công trình Nam Hà Nội"
python -m ketoan_ai.cli hoa-don email --ngay 7
python -m ketoan_ai.cli hoa-don xuat-misa
python -m ketoan_ai.cli doi-chieu --misa samples/bang_ke_misa_mau.xlsx --thue samples/bang_ke_thue_mau.xlsx
python -m ketoan_ai.cli luong --nhan-su samples/nhan_su_mau.xlsx --cham-cong samples/cham_cong_mau.xlsx --danh-gia samples/danh_gia_mau.xlsx --thang 10/2026
python -m pytest -q      # kiểm thử offline
```

## Cấu trúc

```
config/                 Cấu hình & danh mục (SỬA Ở ĐÂY cho khớp MISA của công ty)
  settings.yaml           tài khoản ngân hàng, quy tắc hạch toán, email, TK mặc định
  khoan_muc_chi_phi.yaml  danh mục khoản mục chi phí + từ khoá
  cong_trinh.yaml         danh mục công trình + bí danh
  doi_tuong.yaml          nhà cung cấp / khách hàng (mã MISA, MST)
  misa_templates.yaml     tên cột file nhập khẩu MISA
  luong_rules.yaml        quy tắc tính lương
ketoan_ai/
  ai.py                   gọi Claude (đọc PDF/ảnh, phân loại) - có fallback theo từ khoá
  bank/                   sao kê: parser, Google Sheet, xuất MISA
  invoices/               hóa đơn: XML TT78, Gmail IMAP, mở link (Playwright), kho SQLite, xuất MISA
  reconcile/              đối chiếu bảng kê
  payroll/                tính lương
  cli.py                  dòng lệnh
app/streamlit_app.py    giao diện web cho kế toán
samples/                dữ liệu mẫu (mô phỏng ảnh chụp thực tế)
data/                   nơi app ghi file nhập MISA, báo cáo, CSDL (tự tạo)
```
