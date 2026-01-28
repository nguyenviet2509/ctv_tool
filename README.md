# 🧮 Công cụ Xử lý Dữ liệu Cộng Tác Viên

Công cụ web giúp tích lũy và quản lý dữ liệu cộng tác viên (CTV) theo tháng từ các file Excel.

## ✨ Tính năng

- ✅ Đọc và xử lý file Excel (.xlsx, .xls)
- ✅ Tích lũy dữ liệu theo CCCD/ID của CTV
- ✅ Tự động cộng dồn tiền hoa hồng, thuế TNCN, số tiền trả CTV
- ✅ Thêm mới CTV khi phát hiện CCCD mới
- ✅ Hiển thị dữ liệu trực quan trong bảng
- ✅ Xuất file Excel tổng hợp
- ✅ Giao diện đẹp, dễ sử dụng
- ✅ Chạy hoàn toàn trên trình duyệt (không cần server)

## 📋 Cấu trúc File Excel

File Excel cần có các cột sau:

| Cột | Tên cột | Mô tả |
|-----|---------|-------|
| A | STT | Số thứ tự |
| B | Tên | Tên cộng tác viên |
| C | SĐT | Số điện thoại |
| F | CCCD/ID | Số CCCD hoặc ID (dùng để định danh CTV) |
| I | Tiền Hoa Hồng | Số tiền hoa hồng |
| J | Thuế TNCN | Thuế thu nhập cá nhân |
| K | Số Tiền Trả CTV | Số tiền thực trả cho CTV |

## 🚀 Cách sử dụng

### Bước 1: Mở công cụ
- Mở file `index.html` bằng trình duyệt web (Chrome, Edge, Firefox, Safari...)
- Hoặc double-click vào file `index.html`

### Bước 2: Upload file mẫu (Tháng đầu tiên)
1. Click vào ô "Upload File Mẫu (Tháng đầu tiên)"
2. Chọn file Excel tháng 1 (hoặc tháng bắt đầu)
3. Hệ thống sẽ đọc và lưu dữ liệu làm file gốc

### Bước 3: Upload các file tháng tiếp theo
1. Click vào ô "Upload File Tháng Tiếp Theo"
2. Chọn file Excel tháng 2, 3, 4... lần lượt
3. Hệ thống sẽ tự động:
   - **Nếu CCCD đã tồn tại**: Cộng dồn các giá trị (Hoa hồng, Thuế, Tiền trả)
   - **Nếu CCCD mới**: Thêm CTV mới vào danh sách

### Bước 4: Xem kết quả
- Dữ liệu tổng hợp được hiển thị trong bảng
- Thông tin thống kê hiển thị ở phần "Thông tin tổng hợp"

### Bước 5: Xuất file Excel
1. Click nút "Xuất File Excel"
2. File sẽ được tải về với tên: `DuLieu_CTV_TongHop_YYYYMMDD.xlsx`

### Bước 6: Bắt đầu lại (nếu cần)
- Click nút "Bắt đầu lại" để xóa toàn bộ dữ liệu và upload file mới

## 💡 Lưu ý quan trọng

1. **Định dạng file**: Chỉ hỗ trợ file `.xlsx` và `.xls`
2. **Cột CCCD/ID**: Đây là cột quan trọng để định danh CTV, phải nằm ở **cột F**
3. **Thứ tự upload**: Bắt buộc phải upload file mẫu trước khi upload các file tháng
4. **Dữ liệu số**: Các cột tiền (I, J, K) phải chứa giá trị số hoặc để trống
5. **Dòng tiêu đề**: Dòng đầu tiên trong file Excel là tiêu đề, sẽ được giữ nguyên

## 🔧 Xử lý logic

### Case 1: CCCD đã tồn tại
```
Tiền Hoa Hồng (mới) = Tiền Hoa Hồng (cũ) + Tiền Hoa Hồng (file tháng)
Thuế TNCN (mới) = Thuế TNCN (cũ) + Thuế TNCN (file tháng)
Số Tiền Trả (mới) = Số Tiền Trả (cũ) + Số Tiền Trả (file tháng)
```

### Case 2: CCCD mới
```
Thêm hàng mới vào cuối danh sách với đầy đủ thông tin từ file tháng
```

## 📁 Cấu trúc dự án

```
tool_ctv/
├── index.html      # Giao diện chính
├── styles.css      # File CSS cho giao diện
├── app.js          # Logic xử lý JavaScript
└── README.md       # File hướng dẫn này
```

## 🌐 Yêu cầu hệ thống

- Trình duyệt web hiện đại (Chrome 90+, Edge 90+, Firefox 88+, Safari 14+)
- JavaScript phải được bật
- Không cần kết nối Internet sau khi tải trang (sử dụng CDN cho thư viện SheetJS)

## 📚 Thư viện sử dụng

- **SheetJS (xlsx)**: Thư viện đọc/ghi file Excel
  - Phiên bản: 0.20.1
  - Link: https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js

## 🐛 Xử lý lỗi

Nếu gặp lỗi:
1. Kiểm tra file Excel có đúng định dạng không
2. Kiểm tra các cột (F, I, J, K) có đúng vị trí không
3. Kiểm tra dữ liệu số có hợp lệ không
4. Mở Console của trình duyệt (F12) để xem chi tiết lỗi

## 📞 Hỗ trợ

Nếu gặp vấn đề khi sử dụng, vui lòng kiểm tra:
- File Excel có đúng cấu trúc
- CCCD/ID phải nằm ở cột F
- Các cột số liệu (I, J, K) phải chứa số

## 📝 Ví dụ

**File Tháng 1:**
| STT | Tên | SĐT | ... | CCCD | ... | Hoa Hồng | Thuế | Tiền Trả |
|-----|-----|-----|-----|------|-----|----------|------|----------|
| 1 | Nguyễn Văn A | 0901234567 | ... | 001234567890 | ... | 5000000 | 500000 | 4500000 |

**File Tháng 2:**
| STT | Tên | SĐT | ... | CCCD | ... | Hoa Hồng | Thuế | Tiền Trả |
|-----|-----|-----|-----|------|-----|----------|------|----------|
| 1 | Nguyễn Văn A | 0901234567 | ... | 001234567890 | ... | 3000000 | 300000 | 2700000 |
| 2 | Trần Thị B | 0987654321 | ... | 002345678901 | ... | 4000000 | 400000 | 3600000 |

**Kết quả sau khi xử lý:**
| STT | Tên | SĐT | ... | CCCD | ... | Hoa Hồng | Thuế | Tiền Trả |
|-----|-----|-----|-----|------|-----|----------|------|----------|
| 1 | Nguyễn Văn A | 0901234567 | ... | 001234567890 | ... | 8000000 | 800000 | 7200000 |
| 2 | Trần Thị B | 0987654321 | ... | 002345678901 | ... | 4000000 | 400000 | 3600000 |

---

**© 2026 - Công cụ xử lý dữ liệu CTV**
