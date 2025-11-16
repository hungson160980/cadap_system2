# ai/gemini_client.py
"""Module tích hợp Google Gemini API"""

import google.generativeai as genai
from typing import List, Dict, Optional
import streamlit as st


class GeminiClient:
    """Client để tương tác với Gemini API"""
    
    def __init__(self, api_key: str, model_name: str = "gemini-2.0-flash-exp"):
        """
        Khởi tạo Gemini client
        
        Args:
            api_key: API key của Google Gemini
            model_name: Tên model (mặc định: gemini-2.0-flash-exp)
        """
        self.api_key = api_key
        self.model_name = model_name
        
        # Cấu hình API
        genai.configure(api_key=api_key)
        
        # Khởi tạo model
        self.model = genai.GenerativeModel(model_name)
        
        # Cấu hình generation
        self.generation_config = {
            'temperature': 0.7,
            'top_p': 0.95,
            'top_k': 40,
            'max_output_tokens': 8192,
        }
    
    def analyze_from_file(self, file_content: str) -> str:
        """
        Phân tích phương án từ nội dung file
        
        Args:
            file_content: Nội dung file đã upload
            
        Returns:
            Kết quả phân tích
        """
        prompt = f"""
Bạn là chuyên gia thẩm định tín dụng ngân hàng. Hãy phân tích phương án sử dụng vốn sau đây một cách chuyên sâu và đưa ra đánh giá:

NỘI DUNG PHƯƠNG ÁN:
{file_content}

YÊU CẦU PHÂN TÍCH:
1. Đánh giá tính khả thi của phương án
2. Phân tích các rủi ro tiềm ẩn
3. Đánh giá năng lực tài chính và khả năng trả nợ
4. Phân tích tài sản bảo đảm (nếu có)
5. Đưa ra các khuyến nghị cụ thể cho cán bộ tín dụng

Hãy trả lời một cách ngắn gọn, chuyên nghiệp nhưng đầy đủ các khía cạnh quan trọng.
"""
        
        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config
            )
            return response.text
        except Exception as e:
            return f"Lỗi khi phân tích: {str(e)}"
    
    def analyze_from_data(self, data: Dict) -> str:
        """
        Phân tích phương án từ dữ liệu đã nhập/chỉnh sửa
        
        Args:
            data: Dictionary chứa dữ liệu đã nhập
            
        Returns:
            Kết quả phân tích
        """
        prompt = f"""
Bạn là chuyên gia thẩm định tín dụng ngân hàng. Hãy phân tích các chỉ số tài chính sau đây và đưa ra đánh giá chuyên sâu:

THÔNG TIN KHÁCH HÀNG:
- Họ tên: {data.get('customer_name', 'N/A')}
- CCCD: {data.get('customer_cccd', 'N/A')}
- Địa chỉ: {data.get('customer_address', 'N/A')}

THÔNG TIN KHOẢN VAY:
- Mục đích vay: {data.get('loan_purpose', 'N/A')}
- Số tiền vay: {data.get('loan_amount', 0):,.0f} VND
- Lãi suất: {data.get('interest_rate', 0)}%/năm
- Thời hạn: {data.get('loan_term', 0)} tháng
- Trả nợ hàng tháng: {data.get('monthly_payment', 0):,.0f} VND

THÔNG TIN TÀI CHÍNH:
- Thu nhập tháng: {data.get('monthly_income', 0):,.0f} VND
- Chi phí tháng: {data.get('monthly_expense', 0):,.0f} VND
- Dòng tiền ròng: {data.get('net_cash_flow', 0):,.0f} VND
- DSR: {data.get('dsr', 0):.2f}%
- Biên an toàn: {data.get('safety_margin', 0):.2f}%

TÀI SẢN BẢO ĐẢM:
- Loại tài sản: {data.get('collateral_type', 'N/A')}
- Giá trị: {data.get('collateral_value', 0):,.0f} VND
- LTV: {data.get('ltv', 0):.2f}%

YÊU CẦU PHÂN TÍCH:
1. So sánh với kết quả phân tích từ file gốc (nếu có sự khác biệt)
2. Đánh giá các chỉ số tài chính hiện tại
3. Phân tích rủi ro dựa trên DSR, LTV, dòng tiền
4. Đưa ra khuyến nghị cho cán bộ tín dụng về việc:
   - Chấp thuận/từ chối khoản vay
   - Điều kiện bổ sung (nếu cần)
   - Các biện pháp giảm thiểu rủi ro

Hãy trả lời một cách ngắn gọn nhưng chuyên sâu, tập trung vào các điểm quan trọng.
"""
        
        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config
            )
            return response.text
        except Exception as e:
            return f"Lỗi khi phân tích: {str(e)}"
    
    def chat(self, message: str, chat_history: Optional[List[Dict]] = None) -> str:
        """
        Chat với Gemini
        
        Args:
            message: Tin nhắn từ người dùng
            chat_history: Lịch sử chat (optional)
            
        Returns:
            Phản hồi từ Gemini
        """
        try:
            # Nếu có lịch sử chat, tạo context
            if chat_history and len(chat_history) > 0:
                context = "Lịch sử hội thoại:\n"
                for msg in chat_history[-5:]:  # Lấy 5 tin nhắn gần nhất
                    context += f"{msg['role']}: {msg['content']}\n"
                context += f"\nTin nhắn hiện tại: {message}"
                prompt = context
            else:
                prompt = message
            
            # Thêm system prompt
            full_prompt = f"""
Bạn là trợ lý AI chuyên về thẩm định tín dụng và phân tích tài chính ngân hàng.
Hãy trả lời câu hỏi một cách chuyên nghiệp, chính xác và hữu ích.

{prompt}
"""
            
            response = self.model.generate_content(
                full_prompt,
                generation_config=self.generation_config
            )
            return response.text
        except Exception as e:
            return f"Lỗi: {str(e)}"
    
    def generate_report_summary(self, data: Dict) -> str:
        """
        Tạo tóm tắt báo cáo thẩm định
        
        Args:
            data: Dữ liệu phương án
            
        Returns:
            Tóm tắt báo cáo
        """
        prompt = f"""
Hãy tạo một tóm tắt báo cáo thẩm định ngắn gọn (2-3 đoạn văn) cho phương án sau:

Khách hàng: {data.get('customer_name', 'N/A')}
Số tiền vay: {data.get('loan_amount', 0):,.0f} VND
Mục đích: {data.get('loan_purpose', 'N/A')}
DSR: {data.get('dsr', 0):.2f}%
Dòng tiền ròng: {data.get('net_cash_flow', 0):,.0f} VND
Đánh giá: {data.get('assessment', 'N/A')}

Tóm tắt nên bao gồm:
1. Thông tin tổng quan về khoản vay
2. Đánh giá năng lực trả nợ
3. Kết luận và khuyến nghị
"""
        
        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config
            )
            return response.text
        except Exception as e:
            return f"Không thể tạo tóm tắt: {str(e)}"


@st.cache_resource
def get_gemini_client(api_key: str) -> Optional[GeminiClient]:
    """
    Lấy Gemini client với caching
    
    Args:
        api_key: API key
        
    Returns:
        GeminiClient instance hoặc None nếu lỗi
    """
    if not api_key or api_key.strip() == "":
        return None
    
    try:
        return GeminiClient(api_key)
    except Exception as e:
        st.error(f"Lỗi khởi tạo Gemini: {str(e)}")
        return None
