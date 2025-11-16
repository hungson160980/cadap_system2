# src/utils.py
"""Các hàm tiện ích chung"""

import re
from typing import Union


def format_number(number: Union[int, float], decimals: int = 0) -> str:
    """
    Format số với dấu phân cách hàng nghìn là dấu chấm
    
    Args:
        number: Số cần format
        decimals: Số chữ số thập phân
        
    Returns:
        Chuỗi số đã được format
    """
    if number is None:
        return "0"
    
    try:
        if decimals > 0:
            formatted = f"{float(number):,.{decimals}f}"
        else:
            formatted = f"{int(number):,}"
        
        # Thay dấu phẩy thành dấu chấm cho hàng nghìn
        # và dấu chấm thành dấu phẩy cho thập phân
        formatted = formatted.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
        return formatted
    except (ValueError, TypeError):
        return str(number)


def parse_number(text: str) -> float:
    """
    Parse chuỗi số về float
    
    Args:
        text: Chuỗi số
        
    Returns:
        Số dạng float
    """
    if not text or text == "":
        return 0.0
    
    try:
        # Loại bỏ dấu chấm (phân cách hàng nghìn)
        # Thay dấu phẩy (phân cách thập phân) thành dấu chấm
        cleaned = str(text).replace(".", "").replace(",", ".")
        return float(cleaned)
    except (ValueError, AttributeError):
        return 0.0


def clean_text(text: str) -> str:
    """Làm sạch text"""
    if not text:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()


def validate_phone(phone: str) -> bool:
    """Kiểm tra số điện thoại hợp lệ"""
    if not phone:
        return False
    # Kiểm tra số điện thoại Việt Nam (10-11 số)
    pattern = r'^(0|\+84)[0-9]{9,10}$'
    return bool(re.match(pattern, str(phone).replace(" ", "").replace("-", "")))


def validate_cccd(cccd: str) -> bool:
    """Kiểm tra CCCD/CMND hợp lệ"""
    if not cccd:
        return False
    # CCCD: 12 số, CMND: 9 hoặc 12 số
    cleaned = str(cccd).replace(" ", "").replace("-", "")
    return len(cleaned) in [9, 12] and cleaned.isdigit()


def safe_divide(numerator: float, denominator: float, default: float = 0.0) -> float:
    """Chia an toàn, tránh chia cho 0"""
    try:
        if denominator == 0:
            return default
        return numerator / denominator
    except (ZeroDivisionError, TypeError):
        return default
