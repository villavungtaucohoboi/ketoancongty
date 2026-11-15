# Hướng dẫn triển khai "Trợ lý kế toán AI"

## 0. Cách app hoạt động (đọc trước)

MISA SME.NET là phần mềm cài trên máy (desktop) và **không có API** cho bên ngoài ghi dữ liệu. Vì vậy app
**không** tự gõ vào màn hình MISA (cách này rất dễ hỏng mỗi khi MISA đổi giao diện). Thay vào đó app dùng chính
chức năng chính thức của MISA: **Tệp → Nhập khẩu từ Excel**.

Luồng chuẩn cho mỗi phân hệ:

1. App lấy dữ liệu (Google Sheet, email, file).
2. AI đọc + phân loại, tô vàng những dòng chưa chắc.
3. Kế toán xem lại 1 phút trên giao diện, sửa nếu cần, bấm xuất file.
4. Mở MISA → Nhập khẩu từ Excel → chọn file → *Ghép cột* (lần đầu, MISA nhớ cho các lần sau) → Nhập.

> Nếu công ty đang dùng **MISA AMIS Kế toán** (bản web) thì AMIS có API mở, có thể đẩy thẳng chứng từ
> không cần bước nhập Excel. Cho tôi biết nếu đúng vậy để nâng cấp.

AI (Claude) được dùng ở 3 chỗ: đọc hóa đơn PDF/ảnh, phân loại khoản mục chi phí + công trình, phân loại
giao dịch ngân hàng. Hóa đơn dạng **XML** được đọc trực tiếp (chính xác 100%, không tốn phí AI).
Không có API key app vẫn chạy nhưng chỉ phân loại theo từ khoá.

---

## 1. Cài đặt trên máy kế toán (Windows)

1. Cài Python 3.11+ từ python.org (tick "Add Python to PATH").
2. Giải nén thư mục `ketoan-ai` vào `D:\ketoan-ai` (hoặc nơi khác).
3. Mở Command Prompt trong thư mục đó, chạy:

```bash
pip install -r requirements.txt
```

4. Chỉ khi cần mở link tra cứu hóa đơn (email không đính kèm file):

```bash
playwright install chromium
```

5. Sao chép `.env.example` thành `.env` và điền (xem mục 2–4).
6. Chạy app:

```bash
streamlit run app/streamlit_app.py
```

Trình duyệt mở `http://localhost:8501`. Có thể tạo file `chay_app.bat` với nội dung trên để bấm đúp là chạy.

---

## 2. Khóa API Claude

1. Vào https://console.anthropic.com → API Keys → Create Key.
2. Dán vào `.env`: `ANTHROPIC_API_KEY=sk-ant-...`
3. Nạp tiền trả trước (prepaid). Chi phí ước tính: đọc 1 hóa đơn PDF ≈ 100–300 đ, phân loại 1 ngày sao kê ≈ 100–500 đ.
   Một tháng vài trăm hóa đơn + sao kê hàng ngày thường **dưới 200.000 đ**.

Model mặc định `claude-opus-5-5`. Muốn rẻ hơn (kém chính xác hơn) đặt `KT_MODEL=claude-sonnet-5-5`.

---

## 3. Phân hệ 1 – Sao kê ngân hàng

### 3.1 Kết nối Google Sheet "Sao kê ngân hàng 2026"

1. Vào https://console.cloud.google.com → tạo project → *APIs & Services* → bật **Google Sheets API** và **Google Drive API**.
2. *Credentials* → *Create credentials* → **Service account** → tạo xong vào tab *Keys* → *Add key* → JSON. Tải file về,
   đổi tên `service_account.json`, đặt vào thư mục `config/`.
3. Mở file JSON, chép địa chỉ `client_email` (dạng `xxx@yyy.iam.gserviceaccount.com`).
4. Mở Google Sheet sao kê → *Chia sẻ* → dán email đó → quyền *Người xem*.
5. Lấy ID sheet từ URL (`https://docs.google.com/spreadsheets/d/<ID>/edit`) → `.env`: `KT_SHEET_ID=<ID>`.

Cách khác không cần Google Cloud: kế toán tải file sao kê (xlsx/csv) từ ngân hàng rồi upload trực tiếp trên tab 1.

### 3.2 Khai báo tài khoản ngân hàng

Sửa `config/settings.yaml` → `sao_ke.tabs`: mỗi tab trên sheet (ACB_CTY, VIETIN_CTY, BIDV_CTY, BIDV_NOCB...) ghi
**số tài khoản đúng như đã khai trong MISA** (Danh mục → Tài khoản ngân hàng) và TK hạch toán (1121/11211...).

### 3.3 Quy tắc hạch toán

`sao_ke.quy_tac` trong `settings.yaml`: các từ khoá chắc chắn (phí ngân hàng → 6428, lãi tiền gửi → 515, nộp thuế → 3331,
lương → 334...). Các giao dịch còn lại AI phân loại; đối tượng/công trình được tra theo `doi_tuong.yaml` và `cong_trinh.yaml`.
Càng khai nhiều bí danh (cách viết tắt hay gặp trong nội dung chuyển khoản, ví dụ "CT NAM HA NOI") thì càng chính xác.

### 3.4 Nhập vào MISA

Tab 1 → chọn ngân hàng, khoảng ngày → *Đọc & phân loại* → xem bảng (dòng vàng = cần kiểm tra) → tải file.
File có 2 sheet: *Thu tien gui* và *Chi tien gui*. Trong MISA: Tệp → Nhập khẩu từ Excel → chọn loại chứng từ
tương ứng (Thu tiền gửi / Chi tiền gửi hoặc "Thu, chi tiền gửi" tuỳ phiên bản) → chọn file và sheet → ghép cột → Nhập.

---

## 4. Phân hệ 2 – Hóa đơn mua vào từ email

### 4.1 Email chung

Gmail: Tài khoản Google → Bảo mật → bật *Xác minh 2 bước* → *Mật khẩu ứng dụng* → tạo cho "Mail" → chép 16 ký tự.
Điền `.env`:

```
KT_EMAIL_USER=hoadon@congty.vn
KT_EMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
```

Email công ty khác Gmail (Outlook, Zoho...): đổi `KT_EMAIL_IMAP_HOST`.

### 4.2 Quy ước với các bộ phận (QUAN TRỌNG)

Khi forward hóa đơn về email chung, **dòng đầu tiên** của email ghi **tên công trình** (như ảnh "Xưởng in Minh Đức").
App lấy dòng này làm công trình. Tên phải nằm trong `config/cong_trinh.yaml` (hoặc một bí danh của nó).

### 4.3 Hai dạng email

- **Có file đính kèm** (zip/xml/pdf): app ưu tiên XML (chính xác tuyệt đối), nếu chỉ có PDF thì AI đọc PDF.
- **Chỉ có link tra cứu**: app mở link bằng trình duyệt ẩn, tải PDF/XML nếu có nút tải, nếu không thì chụp ảnh trang
  và AI đọc ảnh. Mỗi nhà cung cấp HĐĐT (EasyInvoice, M-Invoice, VNPT, Viettel, BKAV...) có trang khác nhau; nếu link nào
  không đọc được, app báo cảnh báo để kế toán tải tay rồi upload trên tab 2.

### 4.4 Khoản mục chi phí & công trình

- `config/khoan_muc_chi_phi.yaml`: tên khoản mục **phải trùng** với Danh mục → Khoản mục chi phí trong MISA.
  Danh sách hiện có: NVL bê tông, thép, gạch, cát, xi măng, NVL khác, nhân công, máy thi công, vận chuyển,
  khảo sát/thiết kế, văn phòng, sửa chữa, khác. Thêm bớt tuỳ ý, kèm từ khoá gợi ý.
- `config/doi_tuong.yaml`: nhà cung cấp (mã MISA + MST). NCC chưa có trong danh mục sẽ được đặt mã tạm
  `NCC_<6 số cuối MST>` và tô vàng để kế toán thêm vào MISA trước khi nhập.

### 4.5 Quy trình hàng ngày

Tab 2 → *Quét email hóa đơn* → hàng đợi hiện ra → sửa khoản mục/công trình nếu AI đoán sai → *Duyệt & xuất file MISA*.
Trong MISA: Tệp → Nhập khẩu từ Excel → *Mua hàng* (hóa đơn mua hàng hóa/dịch vụ) → chọn file → ghép cột → Nhập.
Hóa đơn đã xử lý được lưu trong `data/ketoan.sqlite` để không nhập trùng; email đã xử lý được gắn nhãn `AI_Da_Nhap`.

Có thể đặt lịch tự quét email mỗi giờ bằng Task Scheduler của Windows với lệnh:
`python -m ketoan_ai.cli hoa-don email --ngay 2`

---

## 5. Phân hệ 3 – Đối chiếu bảng kê (ngày 10–12 hàng tháng)

Cổng hoadondientu.gdt.gov.vn yêu cầu đăng nhập + mã captcha nên **kế toán vẫn tải tay** như hiện nay:
Tra cứu → Tra cứu hóa đơn → Hóa đơn điện tử mua vào → chọn kỳ → *Xuất Excel*.
Trên MISA: Báo cáo → Thuế → *Bảng kê hóa đơn, chứng từ hàng hóa dịch vụ mua vào* → xuất Excel.

Tab 3 → upload 2 file → *Đối chiếu*. Kết quả 4 nhóm: Khớp / Lệch tiền / Chỉ có trên Thuế (chưa nhập MISA) /
Chỉ có trên MISA. Tải báo cáo Excel để lưu hồ sơ. App tự nhận diện tiêu đề cột nên không cần sửa file.

---

## 6. Phân hệ 4 – Tính lương: THÔNG TIN CẦN CUNG CẤP

App đã có engine tính lương mẫu chạy được với file mẫu trong `samples/`. Để chạy đúng cho công ty, cần gửi thêm:

1. **Bảng lương mẫu** của 1 tháng gần nhất (Excel, có công thức càng tốt) – để lắp đúng cột và công thức.
2. **Bảng chấm công / ngày công** mẫu và **bảng đánh giá công việc** mẫu (file đang cập nhật hàng tháng).
3. **Bảng nhân sự**: mã NV, lương cơ bản/thoả thuận, lương đóng BHXH, các khoản phụ cấp (ăn trưa, điện thoại, xăng xe,
   trách nhiệm...), số người phụ thuộc đã đăng ký, có đóng BH hay không, loại hợp đồng.
4. **Quy chế**: số ngày công chuẩn (26 hay theo tháng), cách tính lương ngày (lương/26 hay lương/ngày công thực tế),
   hệ số tăng ca, quy đổi xếp loại A/B/C → thưởng/phạt, phép năm, đi muộn, tạm ứng, khấu trừ khác.
5. Nhân viên công trường có tính theo **khoán/ngày công công trình** không, lương có phân bổ theo công trình không
   (để xuất đúng chứng từ lương vào MISA theo công trình).
6. Có cần xuất **bảng lương nhập MISA** (phân hệ Tiền lương) và **bảng kê nộp BHXH / thuế TNCN** hay chỉ cần bảng lương Excel.

Các thông số pháp lý (tỷ lệ BH, trần lương, giảm trừ gia cảnh, biểu thuế) nằm trong `config/luong_rules.yaml`,
mặc định theo quy định phổ biến; **kế toán kiểm tra lại mức áp dụng cho kỳ hiện hành** trước khi dùng.

---

## 7. Bảo mật & vận hành

- `.env` chứa API key và mật khẩu email: không gửi cho ai, không chép lên Drive chung.
- Hóa đơn, sao kê, CSDL nằm trong `data/` trên máy kế toán; chỉ nội dung hóa đơn/giao dịch được gửi lên Claude để đọc.
- Nên backup thư mục `data/` cùng với backup MISA.
- Luôn có bước **kế toán duyệt** trước khi nhập MISA; AI chỉ đề xuất.

## 8. Khắc phục sự cố

| Lỗi | Cách xử lý |
|-----|-----------|
| "Không tìm thấy dòng tiêu đề sao kê" | Sao kê thiếu cột Ngày/Nội dung/Tiền ra/Tiền vào; thêm tên cột vào `COL_ALIASES` trong `ketoan_ai/bank/parser.py` |
| "Chưa cấu hình KT_EMAIL_USER" | Điền `.env` |
| IMAP login failed | Dùng Mật khẩu ứng dụng (không phải mật khẩu Gmail thường); bật IMAP trong Cài đặt Gmail |
| MISA báo sai cột khi nhập | Dùng chức năng ghép cột của MISA, hoặc sửa tên cột trong `config/misa_templates.yaml` |
| NCC tô vàng "chưa có trong danh mục" | Thêm NCC vào MISA và vào `config/doi_tuong.yaml` |
