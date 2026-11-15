"""Tạo bộ dữ liệu mẫu để chạy thử app (mô phỏng theo ảnh chụp màn hình thực tế)."""
from pathlib import Path
import zipfile
from openpyxl import Workbook

S = Path(__file__).resolve().parent

# 1. Sao kê ACB (giống tab ACB_CTY trên Google Sheet)
wb = Workbook(); ws = wb.active; ws.title = "ACB_CTY"
ws.append(["Chủ tài khoản: CONG TY CP XAY DUNG DAN DUNG CONG NGHIEP HA NOI"])
ws.append(["Địa chỉ: LK01-81 KDT C2 GAMUDA Phuong Hoang Mai Thanh pho Ha Noi"])
ws.append([])
ws.append(["", "", "", "BẢNG SAO KÊ GIAO DỊCH - Số tài khoản (VND): 3333558866"])
ws.append(["", "", "", "Từ ngày: 01/01/2026 Đến ngày: 12/01/2026"])
ws.append(["", "", "", "Số dư đầu", "(-) Tổng tiền rút ra", "(+) Tổng tiền gửi vào", "Số dư cuối"])
ws.append(["", "", "", "1.839.887.426", "88.740.986.998", "88.151.814.763", "1.250.715.191"])
ws.append(["Ngày hiệu lực", "Ngày giao dịch", "Số GD", "Nội dung giao dịch", "Số tiền rút ra", "Số tiền gửi vào", "Số dư", "Mã Công Trình", "Tháng - Năm"])
rows = [
    ("06/01/2026", "06/01/2026", 4472, "HCIC TT PHI CONG CHUNG THE CHAP CHO VP LUAT GIA KHANH-060126-17:47:21 310385", "11.340.000", "", "1.520.228.489", "", 12026),
    ("08/01/2026", "08/01/2026", 4473, "CTCP DAU TU BAT DONG SAN DONG DOBQP-CHUYEN TIEN T U THEO HDTC SO 3012 NGAY 30.12.2025 VV: KHAO SAT DIA CHAT CONG TRINH KHU TAI DINH CU, GARA XE, CAY XANH TAP TRUNG", "", "32.417.280", "1.552.645.769", "", 12026),
    ("09/01/2026", "09/01/2026", 4474, "HCIC TT CONG NO BE TONG CT NAM HA NOI CHO CTY TIEN PHAT-090126-08:55:14 520237", "300.000.000", "", "1.252.645.769", "", 12026),
    ("09/01/2026", "09/01/2026", 4475, "HCIC TT CONG NO BE TONG CT NAM HA NOI CHO CTY TIEN PHAT-090126-08:55:15 520242", "200.000.000", "", "1.052.645.769", "", 12026),
    ("09/01/2026", "09/01/2026", 4476, "CONG TY CP XAY DUNG DAN DUNG CONG NGHIEP HA NOI##23798 TONG CTY DT PT NHA VA DO THI NAM HA NOI TT TIEN TCXDCTTMDV DONG VAN THEO HD SO 26/2025/HD-XD/NHN-HCIC-0103775584##", "", "4.549.883.808", "5.602.529.577", "", 12026),
    ("09/01/2026", "09/01/2026", 4477, "HCIC TT SUA MAY TINH VP CHO CTY CONG NGHE AN PHAT-090126-16:28:47 307655", "2.106.000", "", "5.600.423.577", "", 12026),
    ("09/01/2026", "09/01/2026", 4478, "HCIC TT CAP CT NAM HA NOI CHO CTY DV AN HUY-090126-16:28:47 307773", "1.926.000", "", "5.598.497.577", "", 12026),
    ("10/01/2026", "10/01/2026", 4480, "HCIC TT TON CT NHA B AN KHANH CHO CTY NGOC DAN-100126-12:16:51 931487", "109.109.000", "", "5.465.388.577", "", 12026),
    ("10/01/2026", "10/01/2026", 4481, "HCIC TT VAT TU CT NAM HA NOI CHO CTY KIM TUAN PHAT-100126-12:16:52 931241", "4.065.000", "", "5.461.323.577", "", 12026),
    ("10/01/2026", "10/01/2026", 4484, "HCIC TT NUOC UONG CHO CTY VIET A-100126-12:16:53 931392", "6.487.000", "", "5.417.494.577", "", 12026),
    ("12/01/2026", "12/01/2026", 4490, "PHI QUAN LY TK THANG 12/2025", "22.000", "", "5.417.472.577", "", 12026),
    ("12/01/2026", "12/01/2026", 4491, "LAI TIEN GUI KHONG KY HAN", "", "1.250.000", "5.418.722.577", "", 12026),
]
for r in rows: ws.append(list(r))
wb.save(S / "sao_ke_ACB_mau.xlsx")

# 2. Hóa đơn XML chuẩn TT78 (giống HĐ số 84 trong ảnh)
xml = """<?xml version="1.0" encoding="UTF-8"?>
<HDon><DLHDon Id="data"><TTChung><PBan>2.0.0</PBan><THDon>HÓA ĐƠN GIÁ TRỊ GIA TĂNG</THDon><KHMSHDon>1</KHMSHDon><KHHDon>C26TMT</KHHDon><SHDon>84</SHDon><NLap>2026-10-29</NLap><DVTTe>VND</DVTTe><TGia>1</TGia><HTTToan>TM/CK</HTTToan></TTChung>
<NDHDon><NBan><Ten>CÔNG TY CỔ PHẦN THƯƠNG MẠI VÀ DỊCH VỤ TÀI NGUYÊN MÔI TRƯỜNG VIỆT NAM</Ten><MST>0108640735</MST><DChi>Số 159 Láng Hạ, Phường Đống Đa, Thành phố Hà Nội, Việt Nam</DChi><SDThoai>024.3835.2004</SDThoai></NBan>
<NMua><Ten>CÔNG TY CỔ PHẦN XÂY DỰNG DÂN DỤNG CÔNG NGHIỆP HÀ NỘI</Ten><MST>0103775584</MST><DChi>Số 6 ngách 259/5 Phố Vọng, Phường Tương Mai, Thành phố Hà Nội, Việt Nam</DChi></NMua>
<DSHHDV><HHDVu><TChat>1</TChat><STT>1</STT><THHDVu>Máy laser 5 tia LSG686SPD</THHDVu><DVTinh>Bộ</DVTinh><SLuong>1</SLuong><DGia>3200000</DGia><ThTien>3200000</ThTien><TSuat>8%</TSuat></HHDVu></DSHHDV>
<TToan><THTTLTSuat><LTSuat><TSuat>8%</TSuat><ThTien>3200000</ThTien><TThue>256000</TThue></LTSuat></THTTLTSuat><TgTCThue>3200000</TgTCThue><TgTThue>256000</TgTThue><TgTTTBSo>3456000</TgTTTBSo><TgTTTBChu>Ba triệu bốn trăm năm mươi sáu nghìn đồng</TgTTTBChu></TToan></NDHDon></DLHDon>
<MCCQT Id="mccqt">000C5D0F05132843FB959160DF96CC7EAF</MCCQT></HDon>"""
(S / "hoa_don_84_mau.xml").write_text(xml, encoding="utf-8")

xml2 = xml.replace("<SHDon>84</SHDon>", "<SHDon>604</SHDon>").replace("C26TMT", "C26TTP").replace("2026-10-29", "2026-10-28") \
    .replace("CÔNG TY CỔ PHẦN THƯƠNG MẠI VÀ DỊCH VỤ TÀI NGUYÊN MÔI TRƯỜNG VIỆT NAM", "CÔNG TY TNHH THƯƠNG MẠI VÀ PHÁT TRIỂN KIM TUẤN PHÁT") \
    .replace("0108640735", "0106789123").replace("Máy laser 5 tia LSG686SPD", "Bê tông thương phẩm M250 đá 1x2").replace("<DVTinh>Bộ</DVTinh>", "<DVTinh>m3</DVTinh>") \
    .replace("<SLuong>1</SLuong><DGia>3200000</DGia>", "<SLuong>40</SLuong><DGia>80000</DGia>").replace("000C5D0F05132843FB959160DF96CC7EAF", "000C5D0F05132843FB959160DF96CC7E00")
(S / "hoa_don_604_betong_mau.xml").write_text(xml2, encoding="utf-8")
with zipfile.ZipFile(S / "hoa_don_604_mau.zip", "w") as z:
    z.writestr("1_C26TTP_604.xml", xml2)

# 3. Email .eml mẫu: bộ phận forward kèm ghi chú tên công trình + zip đính kèm
from email.message import EmailMessage
m = EmailMessage()
m["Subject"] = "Fwd: CÔNG TY TNHH THƯƠNG MẠI VÀ PHÁT TRIỂN KIM TUẤN PHÁT gửi hóa đơn điện tử số 00000604"
m["From"] = "Thu Huong Tran <huong@congty.vn>"
m["To"] = "hoadon@congty.vn"
m["Date"] = "Wed, 29 Oct 2026 17:51:00 +0700"
m.set_content("Công trình Nam Hà Nội\n\n---------- Forwarded message ---------\nTừ: KIM TUAN PHAT <hd@ktp.vn>\n\nKính gửi Quý khách hàng,\nHóa đơn số 604 đính kèm. Quý khách vui lòng truy cập https://tracuu.easyinvoice.com.vn/?code=A9HL5FRU2 để xem.\n")
m.add_attachment((S / "hoa_don_604_mau.zip").read_bytes(), maintype="application", subtype="zip", filename="1_C26TTP_604.zip")
(S / "email_hoa_don_mau.eml").write_bytes(bytes(m))

# 4. Bảng kê MISA & Thuế để đối chiếu
wb = Workbook(); ws = wb.active; ws.title = "BangKe"
ws.append(["BẢNG KÊ HÓA ĐƠN, CHỨNG TỪ HÀNG HÓA, DỊCH VỤ MUA VÀO"])
ws.append(["Từ ngày 01/10/2026 đến ngày 31/10/2026"])
ws.append(["Số hóa đơn", "Ngày hóa đơn", "Ký hiệu", "Tên người bán", "Mã số thuế người bán", "Mặt hàng", "Giá trị HHDV mua vào chưa thuế", "Thuế GTGT", "Tài khoản thuế"])
ws.append(["0000084", "29/10/2026", "1C26TMT", "CÔNG TY CP TM VÀ DV TÀI NGUYÊN MÔI TRƯỜNG VN", "0108640735", "Máy laser 5 tia", 3200000, 256000, "1331"])
ws.append(["0000604", "28/10/2026", "1C26TTP", "CÔNG TY TNHH TM VÀ PT KIM TUẤN PHÁT", "0106789123", "Bê tông M250", 3200000, 256000, "1331"])
ws.append(["0000389", "27/10/2026", "1C26TCD", "CÔNG TY TNHH CHUNG ĐIỆP", "0101234567", "Thép D10", 50000000, 4000000, "1331"])
ws.append(["0000389", "27/10/2026", "1C26TCD", "CÔNG TY TNHH CHUNG ĐIỆP", "0101234567", "Thép D12", 20000000, 1600000, "1331"])
ws.append(["0003561", "25/10/2026", "1C26TMI", "CÔNG TY TNHH VẬT TƯ NGÀNH NƯỚC MINH QUÂN", "0109999888", "Ống nhựa", 12000000, 960000, "1331"])
wb.save(S / "bang_ke_misa_mau.xlsx")

wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
ws.append(["STT", "Ký hiệu mẫu số", "Ký hiệu hóa đơn", "Số hóa đơn", "Ngày lập", "MST người bán", "Tên người bán", "Tổng tiền chưa thuế", "Tổng tiền thuế", "Tổng tiền thanh toán", "Trạng thái hóa đơn"])
ws.append([1, "1", "C26TMT", "84", "29/10/2026", "0108640735", "CÔNG TY CỔ PHẦN THƯƠNG MẠI VÀ DỊCH VỤ TÀI NGUYÊN MÔI TRƯỜNG VIỆT NAM", "3.200.000", "256.000", "3.456.000", "Hóa đơn mới"])
ws.append([2, "1", "C26TTP", "604", "28/10/2026", "0106789123", "CÔNG TY TNHH THƯƠNG MẠI VÀ PHÁT TRIỂN KIM TUẤN PHÁT", "3.200.000", "256.000", "3.456.000", "Hóa đơn mới"])
ws.append([3, "1", "C26TCD", "389", "27/10/2026", "0101234567", "CÔNG TY TNHH CHUNG ĐIỆP", "70.000.000", "5.600.000", "75.600.000", "Hóa đơn mới"])
ws.append([4, "1", "C26TTP", "603", "28/10/2026", "0106789123", "CÔNG TY TNHH THƯƠNG MẠI VÀ PHÁT TRIỂN KIM TUẤN PHÁT", "15.000.000", "1.200.000", "16.200.000", "Hóa đơn mới"])
ws.append([5, "1", "C26TLB", "156", "26/10/2026", "0107777666", "CÔNG TY TNHH DỊCH VỤ KS LABALA", "8.500.000", "680.000", "9.180.000", "Hóa đơn mới"])
wb.save(S / "bang_ke_thue_mau.xlsx")

# 5. Lương: nhân sự, chấm công, đánh giá
wb = Workbook(); ws = wb.active; ws.title = "NhanSu"
ws.append(["Mã NV", "Họ tên", "Bộ phận", "Lương cơ bản", "Lương đóng BH", "Phụ cấp ăn trưa", "Phụ cấp điện thoại", "Phụ cấp xăng xe", "Phụ cấp trách nhiệm", "Số người phụ thuộc", "Tạm ứng"])
ws.append(["NV001", "Nguyễn Văn An", "Kỹ thuật", 18000000, 12000000, 730000, 300000, 500000, 0, 1, 2000000])
ws.append(["NV002", "Trần Thị Bình", "Kế toán", 15000000, 10000000, 730000, 200000, 0, 1000000, 0, 0])
ws.append(["NV003", "Lê Minh Cường", "Chỉ huy CT", 35000000, 20000000, 730000, 500000, 1000000, 3000000, 2, 0])
wb.save(S / "nhan_su_mau.xlsx")
wb = Workbook(); ws = wb.active; ws.title = "ChamCong"
ws.append(["Mã NV", "Họ tên", "Ngày công", "Nghỉ phép", "Tăng ca ngày thường", "Tăng ca ngày nghỉ", "Tăng ca ngày lễ"])
ws.append(["NV001", "Nguyễn Văn An", 24, 2, 10, 8, 0])
ws.append(["NV002", "Trần Thị Bình", 26, 0, 0, 0, 0])
ws.append(["NV003", "Lê Minh Cường", 25, 0, 20, 16, 8])
wb.save(S / "cham_cong_mau.xlsx")
wb = Workbook(); ws = wb.active; ws.title = "DanhGia"
ws.append(["Mã NV", "Họ tên", "Xếp loại", "Thưởng khác", "Phạt"])
ws.append(["NV001", "Nguyễn Văn An", "A", 500000, 0])
ws.append(["NV002", "Trần Thị Bình", "B", 0, 100000])
ws.append(["NV003", "Lê Minh Cường", "A", 2000000, 0])
wb.save(S / "danh_gia_mau.xlsx")
print("Đã tạo dữ liệu mẫu trong", S)
