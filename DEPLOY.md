# 🚀 Hướng Dẫn Deploy Chi Tiết

## Mục Lục
1. [Deploy Local](#1-deploy-local)
2. [Deploy lên Streamlit Cloud](#2-deploy-lên-streamlit-cloud)
3. [Deploy lên Heroku](#3-deploy-lên-heroku-optional)
4. [Troubleshooting](#4-troubleshooting)

---

## 1. Deploy Local

### Bước 1: Chuẩn Bị Môi Trường

```bash
# Kiểm tra Python version (cần >= 3.8)
python --version

# Tạo virtual environment
python -m venv venv

# Kích hoạt virtual environment
# Windows:
venv\Scripts\activate

# Linux/Mac:
source venv/bin/activate
```

### Bước 2: Cài Đặt Dependencies

```bash
# Cài đặt tất cả packages
pip install -r requirements.txt

# Kiểm tra cài đặt
pip list
```

### Bước 3: Chạy Ứng Dụng

```bash
# Chạy Streamlit
streamlit run app.py

# Hoặc chỉ định port cụ thể
streamlit run app.py --server.port 8501
```

### Bước 4: Truy Cập

Mở trình duyệt tại: `http://localhost:8501`

---

## 2. Deploy lên Streamlit Cloud

### Phương Án A: Deploy Qua GitHub (Khuyến Nghị)

#### Bước 1: Tạo GitHub Repository

1. **Tạo repository mới:**
   - Truy cập: https://github.com/new
   - Repository name: `cadap-system` (hoặc tên bạn muốn)
   - Description: "Hệ thống thẩm định phương án kinh doanh"
   - Chọn Public (nếu muốn chia sẻ) hoặc Private
   - **KHÔNG** tích "Add a README file"
   - Click "Create repository"

#### Bước 2: Push Code Lên GitHub

```bash
# Di chuyển vào thư mục dự án
cd cadap_project

# Khởi tạo Git (nếu chưa có)
git init

# Thêm tất cả files
git add .

# Commit
git commit -m "Initial commit: CADAP System"

# Đổi tên branch thành main
git branch -M main

# Thêm remote origin (thay YOUR_USERNAME)
git remote add origin https://github.com/YOUR_USERNAME/cadap-system.git

# Push lên GitHub
git push -u origin main
```

#### Bước 3: Deploy trên Streamlit Cloud

1. **Đăng nhập Streamlit Cloud:**
   - Truy cập: https://share.streamlit.io
   - Click "Sign up" hoặc "Sign in with GitHub"
   - Cho phép Streamlit truy cập GitHub account

2. **Tạo App Mới:**
   - Click "New app" (góc phải trên)
   - Hoặc click "Create app" nếu đây là app đầu tiên

3. **Cấu Hình App:**
   - **Repository:** Chọn `YOUR_USERNAME/cadap-system`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - **App URL (optional):** Tùy chỉnh URL nếu muốn

4. **Advanced Settings (Optional):**
   - Click "Advanced settings"
   - **Python version:** 3.9 hoặc 3.10
   - **Secrets:** Thêm API keys nếu cần (xem bên dưới)

5. **Deploy:**
   - Click "Deploy!"
   - Đợi 2-5 phút để Streamlit build và deploy

#### Bước 4: Cấu Hình Secrets (Optional nhưng Khuyến Nghị)

Để bảo mật API key:

1. Trong Streamlit Cloud, vào app settings
2. Click tab "Secrets"
3. Thêm:
```toml
[secrets]
GEMINI_API_KEY = "your_actual_api_key_here"
```
4. Save

Sau đó update code để đọc từ secrets:
```python
# Trong app.py, thay:
api_key = st.text_input("API Key", type="password")

# Bằng:
import streamlit as st
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.text_input("API Key", type="password")
```

### Phương Án B: Deploy Trực Tiếp (Không Qua GitHub)

1. **Compress dự án:**
```bash
zip -r cadap_project.zip cadap_project/
```

2. **Upload trực tiếp:**
   - Truy cập https://share.streamlit.io
   - Chọn "Deploy from file"
   - Upload file zip
   - Làm theo hướng dẫn

---

## 3. Deploy lên Heroku (Optional)

### Bước 1: Chuẩn Bị Files

Tạo file `Procfile` (không có extension):
```
web: sh setup.sh && streamlit run app.py
```

Tạo file `setup.sh`:
```bash
mkdir -p ~/.streamlit/

echo "\
[server]\n\
headless = true\n\
port = $PORT\n\
enableCORS = false\n\
\n\
" > ~/.streamlit/config.toml
```

Tạo file `runtime.txt`:
```
python-3.9.18
```

### Bước 2: Deploy

```bash
# Đăng nhập Heroku
heroku login

# Tạo app
heroku create cadap-system

# Push code
git push heroku main

# Mở app
heroku open
```

---

## 4. Troubleshooting

### Lỗi: ModuleNotFoundError

**Nguyên nhân:** Thiếu package

**Giải pháp:**
```bash
pip install -r requirements.txt --upgrade
```

### Lỗi: API Key Invalid

**Nguyên nhân:** API key không đúng hoặc hết hạn

**Giải pháp:**
1. Kiểm tra API key tại: https://aistudio.google.com/app/apikey
2. Tạo API key mới nếu cần
3. Cập nhật trong ứng dụng

### Lỗi: File Upload Failed

**Nguyên nhân:** File quá lớn hoặc sai định dạng

**Giải pháp:**
1. Kiểm tra file có đúng định dạng .docx
2. Giảm kích thước file (< 200MB)
3. Kiểm tra cấu hình `maxUploadSize` trong config.toml

### Lỗi: Memory Error on Streamlit Cloud

**Nguyên nhân:** App sử dụng quá nhiều RAM

**Giải pháp:**
1. Tối ưu code (xóa cache không cần thiết)
2. Giảm số lượng biểu đồ hiển thị cùng lúc
3. Upgrade plan Streamlit Cloud (nếu cần)

### Lỗi: Git Push Failed

**Nguyên nhân:** Conflict hoặc sai remote

**Giải pháp:**
```bash
# Kiểm tra remote
git remote -v

# Xóa và thêm lại remote
git remote remove origin
git remote add origin https://github.com/YOUR_USERNAME/cadap-system.git

# Force push (cẩn thận!)
git push -u origin main --force
```

### App Chạy Chậm

**Giải pháp:**
1. Sử dụng `@st.cache_data` và `@st.cache_resource`
2. Giảm số lượng API calls
3. Tối ưu hóa queries và tính toán

---

## 5. Best Practices

### Bảo Mật

1. **Không commit API keys:**
```bash
# Thêm vào .gitignore
.streamlit/secrets.toml
.env
```

2. **Sử dụng Streamlit Secrets:**
```python
# Đọc secrets
api_key = st.secrets["GEMINI_API_KEY"]
```

3. **Validate input:**
```python
if not api_key or len(api_key) < 10:
    st.error("Invalid API key")
    return
```

### Performance

1. **Cache functions:**
```python
@st.cache_data
def load_data():
    return expensive_computation()
```

2. **Lazy loading:**
```python
if st.button("Load data"):
    data = load_expensive_data()
```

3. **Optimize imports:**
```python
# Import only when needed
if chart_type == "matplotlib":
    import matplotlib.pyplot as plt
```

### Monitoring

1. **Add logging:**
```python
import logging
logging.info(f"User uploaded file: {filename}")
```

2. **Track errors:**
```python
try:
    result = risky_operation()
except Exception as e:
    st.error(f"Error: {e}")
    logging.error(f"Error in operation: {e}")
```

3. **Monitor usage:**
- Sử dụng Streamlit Analytics
- Track API usage
- Monitor response times

---

## 6. Update và Maintenance

### Update Code

```bash
# Pull latest changes
git pull origin main

# Make changes
# ... edit files ...

# Commit và push
git add .
git commit -m "Update: Description of changes"
git push origin main
```

Streamlit Cloud sẽ tự động redeploy khi detect thay đổi!

### Update Dependencies

```bash
# Cập nhật package
pip install --upgrade package_name

# Update requirements.txt
pip freeze > requirements.txt

# Commit changes
git add requirements.txt
git commit -m "Update dependencies"
git push
```

### Rollback

```bash
# Xem lịch sử commits
git log

# Rollback về commit cũ
git revert <commit_hash>

# Hoặc reset (cẩn thận!)
git reset --hard <commit_hash>
git push --force
```

---

## 7. Tài Nguyên Hữu Ích

- **Streamlit Docs:** https://docs.streamlit.io
- **Streamlit Community:** https://discuss.streamlit.io
- **Gemini API Docs:** https://ai.google.dev/docs
- **GitHub Guides:** https://guides.github.com

---

## 8. Checklist Deploy

- [ ] Code chạy được local
- [ ] requirements.txt đầy đủ
- [ ] .gitignore đã cấu hình
- [ ] README.md đã viết
- [ ] API keys được bảo mật
- [ ] Push lên GitHub thành công
- [ ] Deploy lên Streamlit Cloud
- [ ] Test app trên cloud
- [ ] Cấu hình secrets (nếu cần)
- [ ] Chia sẻ URL với team

---

**Chúc bạn deploy thành công! 🎉**

Nếu gặp vấn đề, hãy check lại từng bước hoặc tạo Issue trên GitHub.
