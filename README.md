# 🏦 Hệ Thống Thẩm Định Phương Án Kinh Doanh

Ứng dụng web chuyên nghiệp để thẩm định phương án sử dụng vốn của khách hàng, tích hợp AI Gemini để phân tích chuyên sâu.

## ✨ Tính Năng Chính

### 1. 📤 Upload và Trích Xuất Tự Động
- Upload file phương án định dạng `.docx`
- Tự động trích xuất thông tin khách hàng, khoản vay, tài sản bảo đảm
- Cho phép chỉnh sửa dữ liệu thủ công với nút tăng/giảm (+/-)

### 2. 📊 8 Tab Chức Năng Chuyên Nghiệp

#### Tab 1: 👤 Thông Tin Khách Hàng
- Họ tên, CCCD/CMND
- Địa chỉ, số điện thoại
- Kiểm tra tính hợp lệ của dữ liệu

#### Tab 2: 💰 Thông Tin Khoản Vay
- Mục đích vay, tổng nhu cầu vốn
- Vốn đối ứng, số tiền vay
- Lãi suất, thời hạn vay
- Tính toán tự động tỷ lệ vốn đối ứng

#### Tab 3: 🏠 Tài Sản Bảo Đảm
- Loại tài sản, giá trị thị trường
- Địa chỉ tài sản, LTV
- Giấy tờ pháp lý
- Cảnh báo LTV theo mức độ rủi ro

#### Tab 4: 📊 Tính Toán Chỉ Tiêu Tài Chính
- Trả nợ hàng tháng (dư nợ giảm dần)
- DSR (Debt Service Ratio)
- Dòng tiền ròng
- Biên an toàn tài chính
- Đánh giá năng lực trả nợ tự động

#### Tab 5: 📈 Biểu Đồ Trực Quan
- Lịch trả nợ hàng tháng
- Phân tích dòng tiền
- Cơ cấu nguồn vốn
- So sánh thu nhập và nghĩa vụ
- Dư nợ giảm dần theo thời gian

#### Tab 6: 🤖 Phân Tích AI - Gemini
**Phần 1: Phân tích từ File Upload**
- Nguồn: Dữ liệu gốc từ file khách hàng
- Phân tích rủi ro, năng lực trả nợ
- Đề xuất chuyên sâu

**Phần 2: Phân tích từ Dữ Liệu Đã Chỉnh Sửa**
- Nguồn: Dữ liệu sau khi hiệu chỉnh
- So sánh với kết quả file gốc
- Khuyến nghị cho cán bộ tín dụng

#### Tab 7: 💬 Chatbox Gemini
- Hỏi đáp trực tiếp với AI
- Lịch sử hội thoại
- Nút xóa đoạn chat

#### Tab 8: 📥 Xuất File
**Option 1: Xuất Excel**
- Bảng kê kế hoạch trả nợ chi tiết
- Format chuyên nghiệp với màu sắc

**Option 2: Xuất PDF**
- Báo cáo thẩm định đầy đủ
- Bao gồm tất cả thông tin và phân tích AI
- Có thể nhúng biểu đồ

### 3. 🎨 Giao Diện
- Design hiện đại, dễ sử dụng
- Phân cách hàng nghìn bằng dấu "."
- Responsive, tương thích mọi thiết bị
- Màu sắc chuyên nghiệp

### 4. 🔄 Tính Toán Động
- Tự động nhận biết thay đổi dữ liệu
- Tính toán lại tức thời
- Cảnh báo khi dữ liệu không hợp lệ

## 🚀 Hướng Dẫn Cài Đặt

### Yêu Cầu Hệ Thống
- Python 3.8 trở lên
- pip (Python package manager)

### Cài Đặt Local

1. **Clone repository**
```bash
git clone https://github.com/your-username/cadap-system.git
cd cadap-system
```

2. **Tạo virtual environment (khuyến nghị)**
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

3. **Cài đặt dependencies**
```bash
pip install -r requirements.txt
```

4. **Chạy ứng dụng**
```bash
streamlit run app.py
```

5. **Truy cập ứng dụng**
- Mở trình duyệt tại: `http://localhost:8501`

## ☁️ Deploy lên Streamlit Cloud

### Bước 1: Chuẩn Bị Repository

1. **Tạo repository trên GitHub**
   - Truy cập https://github.com
   - Click "New repository"
   - Đặt tên repository (ví dụ: `cadap-system`)
   - Chọn Public hoặc Private
   - Click "Create repository"

2. **Push code lên GitHub**
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/your-username/cadap-system.git
git push -u origin main
```

### Bước 2: Deploy trên Streamlit Cloud

1. **Truy cập Streamlit Cloud**
   - Vào https://share.streamlit.io
   - Đăng nhập bằng GitHub account

2. **Tạo App Mới**
   - Click "New app"
   - Chọn repository: `your-username/cadap-system`
   - Branch: `main`
   - Main file path: `app.py`
   - Click "Deploy"

3. **Đợi Deploy**
   - Streamlit sẽ tự động cài đặt dependencies
   - Quá trình deploy mất khoảng 2-5 phút

4. **Truy Cập App**
   - URL sẽ có dạng: `https://your-app-name.streamlit.app`

## 📖 Hướng Dẫn Sử Dụng

### Bước 1: Cấu Hình API Key
1. Mở sidebar (bên trái)
2. Nhập Gemini API Key vào ô "Gemini API Key"
3. Kiểm tra thông báo "API key đã được cấu hình"

### Bước 2: Upload File
1. Ở sidebar, click "Chọn file phương án (.docx)"
2. Chọn file phương án của khách hàng
3. Click "📊 Phân Tích File"
4. Đợi hệ thống trích xuất dữ liệu

### Bước 3: Kiểm Tra và Chỉnh Sửa Dữ Liệu
1. Xem qua các tab từ 1-3 để kiểm tra thông tin
2. Chỉnh sửa thông tin nếu cần (sử dụng nút +/-)
3. Hệ thống tự động đánh dấu khi có thay đổi

### Bước 4: Tính Toán Chỉ Tiêu
1. Chuyển sang Tab 4 "📊 Chỉ tiêu"
2. Nhập thu nhập, chi phí hàng tháng
3. Xem các chỉ tiêu được tính tự động:
   - Trả nợ hàng tháng
   - DSR
   - Dòng tiền ròng
   - Đánh giá năng lực trả nợ

### Bước 5: Xem Biểu Đồ
1. Chuyển sang Tab 5 "📈 Biểu đồ"
2. Chọn loại biểu đồ muốn xem
3. Phân tích trực quan các chỉ tiêu

### Bước 6: Phân Tích AI
1. Chuyển sang Tab 6 "🤖 AI Phân tích"
2. Click "🔍 Phân Tích File" để phân tích file gốc
3. Click "🔍 Phân Tích Dữ Liệu Hiện Tại" để phân tích sau chỉnh sửa
4. Đọc kết quả phân tích chuyên sâu từ AI

### Bước 7: Chat với AI (Optional)
1. Chuyển sang Tab 7 "💬 Chatbot"
2. Nhập câu hỏi vào ô chat
3. Click "📤 Gửi"
4. Xem phản hồi từ Gemini

### Bước 8: Xuất Báo Cáo
1. Chuyển sang Tab 8 "📥 Xuất file"
2. Chọn loại xuất:
   - Excel: Bảng kê chi tiết trả nợ
   - PDF: Báo cáo thẩm định đầy đủ
3. Click nút tạo file
4. Tải xuống file

## 🔑 Lấy Gemini API Key

1. Truy cập: https://aistudio.google.com/app/apikey
2. Đăng nhập bằng Google Account
3. Click "Create API Key"
4. Copy API key và paste vào ứng dụng

**Lưu ý:** 
- API key miễn phí có giới hạn request
- Không chia sẻ API key với người khác
- Không commit API key lên GitHub

## 📁 Cấu Trúc Dự Án

```
cadap_project/
├── app.py                      # File chính Streamlit
├── requirements.txt            # Dependencies
├── README.md                   # Hướng dẫn
├── .gitignore                 # Git ignore
├── .streamlit/
│   └── config.toml            # Cấu hình Streamlit
├── src/
│   ├── __init__.py
│   ├── config.py              # Cấu hình hệ thống
│   ├── utils.py               # Hàm tiện ích
│   └── docx_parser.py         # Trích xuất DOCX
├── logic/
│   ├── __init__.py
│   └── financial_calculator.py # Tính toán tài chính
├── ai/
│   ├── __init__.py
│   └── gemini_client.py       # Tích hợp Gemini
├── export/
│   ├── __init__.py
│   ├── excel_exporter.py      # Xuất Excel
│   └── pdf_exporter.py        # Xuất PDF
└── ui/
    ├── __init__.py
    └── chart_generator.py     # Vẽ biểu đồ
```

## 🛠️ Công Nghệ Sử Dụng

- **Frontend:** Streamlit
- **Backend:** Python 3.8+
- **AI:** Google Gemini 2.0 Flash
- **Data Processing:** 
  - python-docx (đọc Word)
  - pandas (xử lý dữ liệu)
  - openpyxl (Excel)
- **Visualization:** Matplotlib
- **PDF Generation:** ReportLab

## 📊 Công Thức Tài Chính

### 1. Trả Nợ Hàng Tháng (Dư nợ giảm dần)
```
Trả gốc hàng tháng = Tổng vay / Số tháng
Trả lãi = Dư nợ đầu kỳ × Lãi suất tháng
Tổng trả = Trả gốc + Trả lãi
```

### 2. DSR (Debt Service Ratio)
```
DSR = (Trả nợ hàng tháng + Nợ khác) / Thu nhập × 100%
```

### 3. Dòng Tiền Ròng
```
Dòng tiền = Thu nhập - Chi phí - Trả nợ - Nợ khác
```

### 4. LTV (Loan to Value)
```
LTV = Số tiền vay / Giá trị tài sản × 100%
```

### 5. Biên An Toàn
```
Biên an toàn = Dòng tiền ròng / Thu nhập × 100%
```

## ⚠️ Lưu Ý

1. **API Key:** 
   - Không commit API key lên GitHub
   - Sử dụng Streamlit Secrets cho production

2. **File Upload:**
   - Chỉ chấp nhận file .docx
   - File cần có cấu trúc rõ ràng để trích xuất tốt

3. **Dữ Liệu:**
   - Kiểm tra kỹ dữ liệu sau khi trích xuất
   - Thường xuyên lưu/export báo cáo

4. **Performance:**
   - Mô hình Gemini có thể mất vài giây để phản hồi
   - Tránh spam request để không bị rate limit

## 🤝 Đóng Góp

Mọi đóng góp đều được chào đón! Vui lòng:
1. Fork repository
2. Tạo branch mới (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Mở Pull Request

## 📝 License

Dự án này được phát hành dưới MIT License.

## 👥 Tác Giả

Phát triển bởi Claude AI

## 📞 Hỗ Trợ

Nếu gặp vấn đề, vui lòng:
1. Kiểm tra phần Issues trên GitHub
2. Tạo Issue mới với mô tả chi tiết
3. Email: support@example.com

## 🎯 Roadmap

### Version 2.0 (Kế hoạch)
- [ ] Hỗ trợ nhiều ngôn ngữ
- [ ] Tích hợp database
- [ ] Quản lý nhiều khách hàng
- [ ] Dashboard tổng quan
- [ ] Export Word template
- [ ] Tích hợp thêm AI models

---

**🎉 Cảm ơn bạn đã sử dụng Hệ Thống Thẩm Định Phương Án Kinh Doanh!**
