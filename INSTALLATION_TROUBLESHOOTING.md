# 🔧 XỬ LÝ LỖI CÀI ĐẶT

## ❌ Lỗi: "Preparing metadata (pyproject.toml): still running..."

### Nguyên nhân:
- Package `google-generativeai` hoặc `pandas` đang build từ source
- Mất nhiều thời gian trên máy yếu
- Thiếu build tools

### ✅ Giải pháp:

#### Giải pháp 1: Đợi thêm (Khuyến nghị)
```bash
# Đơn giản là đợi thêm 2-5 phút
# Package đang build, không bị lỗi
```

#### Giải pháp 2: Cài từng package
```bash
# Cài từng cái một để dễ debug
pip install streamlit
pip install python-docx
pip install openpyxl
pip install pandas
pip install matplotlib
pip install reportlab
pip install Pillow

# Cuối cùng mới cài Gemini (package nặng nhất)
pip install google-generativeai
```

#### Giải pháp 3: Dùng requirements-minimal.txt
```bash
pip install -r requirements-minimal.txt
```

#### Giải pháp 4: Upgrade pip và wheel
```bash
pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

#### Giải pháp 5: Sử dụng cache
```bash
pip install -r requirements.txt --use-pep517
```

#### Giải pháp 6: Không dùng binary (Linux)
```bash
pip install -r requirements.txt --no-binary :all:
```

---

## ❌ Lỗi: "ERROR: No matching distribution found"

### Giải pháp:
```bash
# Update pip
pip install --upgrade pip

# Thử lại
pip install -r requirements.txt
```

---

## ❌ Lỗi: "Microsoft Visual C++ 14.0 is required" (Windows)

### Giải pháp:
1. Download và cài đặt Visual C++ Build Tools:
   👉 https://visualstudio.microsoft.com/visual-cpp-build-tools/

2. Hoặc dùng pre-built wheels:
```bash
pip install --only-binary :all: -r requirements.txt
```

---

## ❌ Lỗi: "Command 'gcc' failed" (Linux)

### Giải pháp:
```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install python3-dev build-essential

# CentOS/RHEL
sudo yum install python3-devel gcc

# Sau đó cài lại
pip install -r requirements.txt
```

---

## ❌ Lỗi: "Permission denied"

### Giải pháp:
```bash
# Option 1: Dùng --user
pip install --user -r requirements.txt

# Option 2: Dùng virtual environment (KHUYẾN NGHỊ)
python -m venv venv
source venv/bin/activate  # Linux/Mac
# hoặc venv\Scripts\activate  # Windows
pip install -r requirements.txt
```

---

## ❌ Lỗi: "SSL: CERTIFICATE_VERIFY_FAILED"

### Giải pháp:
```bash
# Option 1: Update certificates
pip install --upgrade certifi

# Option 2: Tạm thời bỏ qua SSL (không an toàn)
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
```

---

## ⚡ CÁCH CÀI ĐẶT NHANH NHẤT

### Cho máy mạnh:
```bash
pip install -r requirements.txt
# Đợi 2-5 phút
```

### Cho máy yếu:
```bash
# Cài từng package, bỏ qua những cái không cần ngay
pip install streamlit python-docx openpyxl matplotlib Pillow
pip install google-generativeai
# pandas và reportlab cài sau nếu cần
```

### Nếu gấp:
```bash
# Chạy app với mock data (không cần AI)
pip install streamlit python-docx openpyxl matplotlib Pillow pandas
# Bỏ qua google-generativeai, sẽ có warning nhưng vẫn chạy được
streamlit run app.py
```

---

## 🐍 KIỂM TRA PHIÊN BẢN PYTHON

```bash
python --version
# Cần >= 3.8

# Nếu có nhiều phiên bản
python3 --version
python3.9 --version

# Dùng phiên bản phù hợp
python3.9 -m pip install -r requirements.txt
```

---

## 📦 KIỂM TRA CÀI ĐẶT

```bash
# Kiểm tra package đã cài
pip list

# Kiểm tra package cụ thể
pip show streamlit
pip show google-generativeai

# Test import
python -c "import streamlit; print('Streamlit OK')"
python -c "import google.generativeai; print('Gemini OK')"
```

---

## 🔄 XÓA VÀ CÀI LẠI

```bash
# Xóa cache pip
pip cache purge

# Xóa virtual environment cũ
rm -rf venv/  # Linux/Mac
# hoặc rmdir /s venv  # Windows

# Tạo lại từ đầu
python -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

---

## 🌐 SỬ DỤNG MIRROR (Nếu ở VN)

```bash
# Aliyun mirror (nhanh hơn ở Châu Á)
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/

# Tsinghua mirror
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

---

## 💡 BEST PRACTICE

### 1. Luôn dùng Virtual Environment
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

### 2. Update pip trước khi cài
```bash
pip install --upgrade pip setuptools wheel
```

### 3. Cài từng nhóm
```bash
# Core packages
pip install streamlit pandas numpy

# Document processing
pip install python-docx openpyxl reportlab

# Visualization
pip install matplotlib Pillow

# AI (cuối cùng)
pip install google-generativeai
```

### 4. Log lỗi để debug
```bash
pip install -r requirements.txt > install.log 2>&1
# Kiểm tra file install.log nếu có lỗi
```

---

## 🆘 VẪN KHÔNG ĐƯỢC?

### Plan B: Dùng Docker (Advanced)

Tạo file `Dockerfile`:
```dockerfile
FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app.py"]
```

Chạy:
```bash
docker build -t cadap-app .
docker run -p 8501:8501 cadap-app
```

### Plan C: Dùng Conda

```bash
conda create -n cadap python=3.9
conda activate cadap
conda install -c conda-forge streamlit pandas matplotlib
pip install python-docx openpyxl google-generativeai reportlab
```

### Plan D: Deploy trực tiếp lên Streamlit Cloud

```bash
# Bỏ qua cài đặt local
# Push code lên GitHub
git init
git add .
git commit -m "Initial commit"
git push

# Deploy tại share.streamlit.io
# Streamlit Cloud sẽ tự cài đặt
```

---

## 📞 HỖ TRỢ

Nếu vẫn gặp vấn đề:
1. Copy full error message
2. Google: "pip [error message]"
3. Check Stack Overflow
4. Create GitHub Issue với log đầy đủ

---

**Good luck! 🍀**
