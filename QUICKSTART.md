# ⚡ QUICKSTART - Hướng Dẫn Nhanh

## 🎯 Chạy Ngay Trong 5 Phút

### Bước 1: Giải Nén (30 giây)
```bash
unzip cadap_project.zip
cd cadap_project
```

### Bước 2: Cài Đặt (2 phút)
```bash
pip install -r requirements.txt
```

### Bước 3: Chạy (10 giây)
```bash
streamlit run app.py
```

### Bước 4: Sử Dụng (2 phút)
1. Mở browser tại `http://localhost:8501`
2. Nhập Gemini API key vào sidebar
3. Upload file .docx
4. Enjoy! 🎉

---

## 🚀 Deploy lên Streamlit Cloud (3 bước)

### 1. Push lên GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/cadap-system.git
git push -u origin main
```

### 2. Deploy
- Vào https://share.streamlit.io
- Click "New app"
- Chọn repository của bạn
- Main file: `app.py`
- Click "Deploy"

### 3. Sử Dụng
App của bạn sẽ online tại: `https://your-app.streamlit.app`

---

## 🔑 Lấy Gemini API Key (1 phút)

1. Vào: https://aistudio.google.com/app/apikey
2. Login bằng Google
3. Click "Create API Key"
4. Copy và paste vào app

---

## ❓ Gặp Lỗi?

### Lỗi: command not found: streamlit
```bash
pip install streamlit
```

### Lỗi: ModuleNotFoundError
```bash
pip install -r requirements.txt --upgrade
```

### Lỗi: API key invalid
- Tạo API key mới tại: https://aistudio.google.com/app/apikey

---

## 📚 Tài Liệu Đầy Đủ

- **README.md** - Hướng dẫn chi tiết
- **DEPLOY.md** - Hướng dẫn deploy chi tiết
- **Code comments** - Giải thích trong code

---

## 🎓 Video Tutorial (Sắp Ra Mắt)

- [ ] Hướng dẫn cài đặt
- [ ] Hướng dẫn sử dụng
- [ ] Hướng dẫn deploy
- [ ] Tùy biến nâng cao

---

## 💡 Tips

1. **Dữ liệu mẫu:** Tạo file .docx với cấu trúc rõ ràng
2. **API Key:** Lưu API key vào Streamlit Secrets
3. **Performance:** Sử dụng cache cho các tính toán nặng
4. **Backup:** Thường xuyên export báo cáo

---

**Happy Coding! 🚀**

Need help? Check README.md hoặc tạo Issue trên GitHub.
