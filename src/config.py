# src/config.py
"""Cấu hình hệ thống"""

# Cấu hình Gemini
GEMINI_MODEL = "gemini-2.0-flash-exp"

# Cấu hình hiển thị
THOUSAND_SEPARATOR = "."
DECIMAL_SEPARATOR = ","

# Cấu hình tài chính
DEFAULT_INTEREST_RATE = 8.5  # %/năm
DEFAULT_LOAN_TERM = 120  # tháng
DEFAULT_LTV = 70  # %

# Màu sắc biểu đồ
CHART_COLORS = {
    'primary': '#1f77b4',
    'secondary': '#ff7f0e',
    'success': '#2ca02c',
    'danger': '#d62728',
    'warning': '#ff9800',
    'info': '#17a2b8'
}

# Văn bản mặc định
DEFAULT_TEXTS = {
    'app_title': '🏦 HỆ THỐNG THẨM ĐỊNH PHƯƠNG ÁN KINH DOANH',
    'sidebar_title': '⚙️ Cấu Hình Hệ Thống',
    'upload_label': '📤 Upload Phương Án Sử Dụng Vốn (.docx)',
}
