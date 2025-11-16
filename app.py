# app.py
"""
Ứng dụng Streamlit - Hệ thống Thẩm Định Phương Án Kinh Doanh
"""

import streamlit as st
import sys
import os

# Thêm thư mục gốc vào Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.config import DEFAULT_TEXTS
from src.utils import format_number, parse_number, validate_phone, validate_cccd
from src.docx_parser import DocxParser
from logic.financial_calculator import FinancialCalculator
from ai.gemini_client import get_gemini_client
from export.excel_exporter import ExcelExporter
from export.pdf_exporter import PDFExporter
from ui.chart_generator import ChartGenerator
import tempfile


# Cấu hình trang
st.set_page_config(
    page_title="Hệ thống Thẩm Định PASDV",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS tùy chỉnh
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4788;
        text-align: center;
        padding: 1rem 0;
        border-bottom: 3px solid #1f4788;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f4788;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.25rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 0.25rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .danger-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 0.25rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    div[data-testid="stNumberInput"] {
        background-color: #ffffff;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        background-color: #f0f2f6;
        border-radius: 4px 4px 0 0;
    }
</style>
""", unsafe_allow_html=True)


def initialize_session_state():
    """Khởi tạo session state"""
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    
    if 'raw_file_content' not in st.session_state:
        st.session_state.raw_file_content = ""
    
    if 'data_modified' not in st.session_state:
        st.session_state.data_modified = False
    
    if 'customer_info' not in st.session_state:
        st.session_state.customer_info = {
            'name': '', 'cccd': '', 'address': '', 'phone': ''
        }
    
    if 'loan_info' not in st.session_state:
        st.session_state.loan_info = {
            'purpose': 'Kinh doanh',
            'total_need': 0.0,
            'equity': 0.0,
            'loan_amount': 0.0,
            'equity_ratio': 0.0,
            'interest_rate': 8.5,
            'loan_term': 120,
            'payment_frequency': 'Tháng'
        }
    
    if 'collateral_info' not in st.session_state:
        st.session_state.collateral_info = {
            'asset_type': 'Bất động sản',
            'market_value': 0.0,
            'asset_address': '',
            'ltv': 70.0,
            'legal_docs': 'Sổ đỏ'
        }
    
    if 'financial_info' not in st.session_state:
        st.session_state.financial_info = {
            'monthly_income': 0.0,
            'monthly_expense': 0.0,
            'other_debt': 0.0
        }
    
    if 'chat_history' not in st.session_state:
        st.session_state.chat_history = []
    
    if 'file_analysis' not in st.session_state:
        st.session_state.file_analysis = ""
    
    if 'data_analysis' not in st.session_state:
        st.session_state.data_analysis = ""


def render_sidebar():
    """Render sidebar"""
    with st.sidebar:
        st.markdown("### ⚙️ Cấu Hình Hệ Thống")
        
        # API Key
        api_key = st.text_input(
            "🔑 Gemini API Key",
            type="password",
            help="Nhập API key của Google Gemini"
        )
        
        st.session_state.api_key = api_key
        
        if api_key:
            st.success("✅ API key đã được cấu hình")
        else:
            st.warning("⚠️ Vui lòng nhập API key để sử dụng tính năng AI")
        
        st.markdown("---")
        
        # Upload file
        st.markdown("### 📤 Upload File")
        uploaded_file = st.file_uploader(
            "Chọn file phương án (.docx)",
            type=['docx'],
            help="Upload file phương án sử dụng vốn của khách hàng"
        )
        
        if uploaded_file is not None:
            if st.button("📊 Phân Tích File", use_container_width=True):
                with st.spinner("Đang trích xuất dữ liệu..."):
                    # Lưu file tạm
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_path = tmp_file.name
                    
                    try:
                        # Parse file
                        parser = DocxParser(tmp_path)
                        parsed_data = parser.parse_full_document()
                        
                        # Cập nhật session state
                        st.session_state.customer_info = parsed_data['customer_info']
                        st.session_state.loan_info = parsed_data['loan_info']
                        st.session_state.collateral_info = parsed_data['collateral_info']
                        st.session_state.financial_info = parsed_data['financial_info']
                        st.session_state.raw_file_content = parsed_data['raw_text']
                        st.session_state.data_loaded = True
                        st.session_state.data_modified = False
                        
                        st.success("✅ Đã trích xuất dữ liệu thành công!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Lỗi khi phân tích file: {str(e)}")
                    finally:
                        # Xóa file tạm
                        if os.path.exists(tmp_path):
                            os.unlink(tmp_path)
        
        st.markdown("---")
        st.markdown("### ℹ️ Hướng Dẫn")
        st.info("""
        1. Nhập API key Gemini
        2. Upload file phương án (.docx)
        3. Xem và chỉnh sửa dữ liệu
        4. Phân tích với AI
        5. Xuất báo cáo
        """)


def render_tab_customer_info():
    """Tab 1: Thông tin định danh khách hàng"""
    st.markdown("### 👤 Thông Tin Định Danh Khách Hàng")
    
    col1, col2 = st.columns(2)
    
    with col1:
        name = st.text_input(
            "Họ và tên",
            value=st.session_state.customer_info['name'],
            key='input_customer_name'
        )
        if name != st.session_state.customer_info['name']:
            st.session_state.customer_info['name'] = name
            st.session_state.data_modified = True
        
        cccd = st.text_input(
            "CCCD/CMND",
            value=st.session_state.customer_info['cccd'],
            key='input_customer_cccd'
        )
        if cccd != st.session_state.customer_info['cccd']:
            st.session_state.customer_info['cccd'] = cccd
            st.session_state.data_modified = True
            
        if cccd and not validate_cccd(cccd):
            st.warning("⚠️ Số CCCD/CMND không hợp lệ (9 hoặc 12 số)")
    
    with col2:
        address = st.text_area(
            "Địa chỉ",
            value=st.session_state.customer_info['address'],
            height=100,
            key='input_customer_address'
        )
        if address != st.session_state.customer_info['address']:
            st.session_state.customer_info['address'] = address
            st.session_state.data_modified = True
        
        phone = st.text_input(
            "Số điện thoại",
            value=st.session_state.customer_info['phone'],
            key='input_customer_phone'
        )
        if phone != st.session_state.customer_info['phone']:
            st.session_state.customer_info['phone'] = phone
            st.session_state.data_modified = True
            
        if phone and not validate_phone(phone):
            st.warning("⚠️ Số điện thoại không hợp lệ")


def render_tab_loan_info():
    """Tab 2: Thông tin tài chính"""
    st.markdown("### 💰 Thông Tin Khoản Vay")
    
    col1, col2 = st.columns(2)
    
    with col1:
        purpose = st.text_input(
            "Mục đích vay",
            value=st.session_state.loan_info['purpose'],
            key='input_loan_purpose'
        )
        if purpose != st.session_state.loan_info['purpose']:
            st.session_state.loan_info['purpose'] = purpose
            st.session_state.data_modified = True
        
        total_need = st.number_input(
            "Tổng nhu cầu vốn (VND)",
            min_value=0.0,
            value=float(st.session_state.loan_info['total_need']),
            step=1000000.0,
            format="%.0f",
            key='input_total_need'
        )
        if total_need != st.session_state.loan_info['total_need']:
            st.session_state.loan_info['total_need'] = total_need
            st.session_state.data_modified = True
        
        equity = st.number_input(
            "Vốn đối ứng (VND)",
            min_value=0.0,
            value=float(st.session_state.loan_info['equity']),
            step=1000000.0,
            format="%.0f",
            key='input_equity'
        )
        if equity != st.session_state.loan_info['equity']:
            st.session_state.loan_info['equity'] = equity
            st.session_state.data_modified = True
        
        loan_amount = st.number_input(
            "Số tiền vay (VND)",
            min_value=0.0,
            value=float(st.session_state.loan_info['loan_amount']),
            step=1000000.0,
            format="%.0f",
            key='input_loan_amount'
        )
        if loan_amount != st.session_state.loan_info['loan_amount']:
            st.session_state.loan_info['loan_amount'] = loan_amount
            st.session_state.data_modified = True
    
    with col2:
        # Tính tỷ lệ vốn đối ứng
        if total_need > 0:
            equity_ratio = (equity / total_need) * 100
            st.session_state.loan_info['equity_ratio'] = equity_ratio
        else:
            equity_ratio = 0
        
        st.metric("Tỷ lệ vốn đối ứng", f"{equity_ratio:.2f}%")
        
        interest_rate = st.number_input(
            "Lãi suất (%/năm)",
            min_value=0.0,
            max_value=100.0,
            value=float(st.session_state.loan_info['interest_rate']),
            step=0.1,
            format="%.2f",
            key='input_interest_rate'
        )
        if interest_rate != st.session_state.loan_info['interest_rate']:
            st.session_state.loan_info['interest_rate'] = interest_rate
            st.session_state.data_modified = True
        
        loan_term = st.number_input(
            "Thời gian vay (tháng)",
            min_value=1,
            max_value=360,
            value=int(st.session_state.loan_info['loan_term']),
            step=12,
            key='input_loan_term'
        )
        if loan_term != st.session_state.loan_info['loan_term']:
            st.session_state.loan_info['loan_term'] = loan_term
            st.session_state.data_modified = True
        
        st.selectbox(
            "Kỳ trả nợ",
            options=['Tháng', 'Quý', 'Năm'],
            index=0,
            key='input_payment_frequency',
            disabled=True
        )
    
    # Kiểm tra logic
    if total_need > 0 and (equity + loan_amount) != total_need:
        st.warning(f"⚠️ Tổng vốn không khớp: {format_number(equity + loan_amount)} ≠ {format_number(total_need)}")


def render_tab_collateral_info():
    """Tab 3: Tài sản bảo đảm"""
    st.markdown("### 🏠 Tài Sản Bảo Đảm")
    
    col1, col2 = st.columns(2)
    
    with col1:
        asset_type = st.text_input(
            "Loại tài sản",
            value=st.session_state.collateral_info['asset_type'],
            key='input_asset_type'
        )
        if asset_type != st.session_state.collateral_info['asset_type']:
            st.session_state.collateral_info['asset_type'] = asset_type
            st.session_state.data_modified = True
        
        market_value = st.number_input(
            "Giá trị thị trường (VND)",
            min_value=0.0,
            value=float(st.session_state.collateral_info['market_value']),
            step=1000000.0,
            format="%.0f",
            key='input_market_value'
        )
        if market_value != st.session_state.collateral_info['market_value']:
            st.session_state.collateral_info['market_value'] = market_value
            st.session_state.data_modified = True
        
        asset_address = st.text_area(
            "Địa chỉ tài sản",
            value=st.session_state.collateral_info['asset_address'],
            height=100,
            key='input_asset_address'
        )
        if asset_address != st.session_state.collateral_info['asset_address']:
            st.session_state.collateral_info['asset_address'] = asset_address
            st.session_state.data_modified = True
    
    with col2:
        # Tính LTV
        if market_value > 0:
            ltv_calculated = (st.session_state.loan_info['loan_amount'] / market_value) * 100
        else:
            ltv_calculated = 0
        
        st.metric("LTV tính toán", f"{ltv_calculated:.2f}%")
        
        ltv = st.number_input(
            "LTV mục tiêu (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(st.session_state.collateral_info['ltv']),
            step=1.0,
            format="%.2f",
            key='input_ltv'
        )
        if ltv != st.session_state.collateral_info['ltv']:
            st.session_state.collateral_info['ltv'] = ltv
            st.session_state.data_modified = True
        
        legal_docs = st.text_area(
            "Giấy tờ pháp lý",
            value=st.session_state.collateral_info['legal_docs'],
            height=100,
            key='input_legal_docs'
        )
        if legal_docs != st.session_state.collateral_info['legal_docs']:
            st.session_state.collateral_info['legal_docs'] = legal_docs
            st.session_state.data_modified = True
    
    # Cảnh báo LTV
    if ltv_calculated > 80:
        st.error(f"🚨 LTV cao ({ltv_calculated:.2f}%) - Rủi ro cao!")
    elif ltv_calculated > 70:
        st.warning(f"⚠️ LTV trung bình ({ltv_calculated:.2f}%) - Cần theo dõi")
    else:
        st.success(f"✅ LTV tốt ({ltv_calculated:.2f}%)")


def render_tab_financial_calculations():
    """Tab 4: Tính toán chỉ tiêu tài chính"""
    st.markdown("### 📊 Tính Toán Chỉ Tiêu Tài Chính")
    
    # Thu nhập - chi phí
    st.markdown("#### Thông tin tài chính")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        monthly_income = st.number_input(
            "Thu nhập tháng (VND)",
            min_value=0.0,
            value=float(st.session_state.financial_info['monthly_income']),
            step=1000000.0,
            format="%.0f",
            key='input_monthly_income'
        )
        if monthly_income != st.session_state.financial_info['monthly_income']:
            st.session_state.financial_info['monthly_income'] = monthly_income
            st.session_state.data_modified = True
    
    with col2:
        monthly_expense = st.number_input(
            "Chi phí tháng (VND)",
            min_value=0.0,
            value=float(st.session_state.financial_info['monthly_expense']),
            step=1000000.0,
            format="%.0f",
            key='input_monthly_expense'
        )
        if monthly_expense != st.session_state.financial_info['monthly_expense']:
            st.session_state.financial_info['monthly_expense'] = monthly_expense
            st.session_state.data_modified = True
    
    with col3:
        other_debt = st.number_input(
            "Nợ khác hàng tháng (VND)",
            min_value=0.0,
            value=float(st.session_state.financial_info['other_debt']),
            step=1000000.0,
            format="%.0f",
            key='input_other_debt'
        )
        if other_debt != st.session_state.financial_info['other_debt']:
            st.session_state.financial_info['other_debt'] = other_debt
            st.session_state.data_modified = True
    
    st.markdown("---")
    
    # Tính toán
    if st.session_state.loan_info['loan_amount'] > 0:
        calc = FinancialCalculator(
            loan_amount=st.session_state.loan_info['loan_amount'],
            interest_rate=st.session_state.loan_info['interest_rate'],
            loan_term=st.session_state.loan_info['loan_term'],
            monthly_income=monthly_income,
            monthly_expense=monthly_expense,
            other_debt=other_debt
        )
        
        summary = calc.get_summary(st.session_state.collateral_info['market_value'])
        
        # Hiển thị các chỉ tiêu
        st.markdown("#### Các chỉ tiêu chính")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "Trả nợ hàng tháng",
                format_number(summary['monthly_payment']) + " VND"
            )
        
        with col2:
            st.metric(
                "Tổng lãi phải trả",
                format_number(summary['total_interest']) + " VND"
            )
        
        with col3:
            dsr = summary['dsr']
            dsr_color = "🟢" if dsr <= 40 else "🟡" if dsr <= 60 else "🔴"
            st.metric(
                f"{dsr_color} DSR",
                f"{dsr:.2f}%" if dsr >= 0 else "N/A"
            )
        
        with col4:
            net_flow = summary['net_cash_flow']
            flow_color = "🟢" if net_flow > 0 else "🔴"
            st.metric(
                f"{flow_color} Dòng tiền ròng",
                format_number(net_flow) + " VND"
            )
        
        st.markdown("---")
        
        # Đánh giá
        st.markdown("#### Đánh giá năng lực trả nợ")
        
        assessment = summary['assessment']
        risk_level = summary['risk_level']
        
        if risk_level == "Thấp":
            st.markdown(f'<div class="success-box"><b>✅ {assessment}</b><br/>Mức độ rủi ro: {risk_level}</div>', 
                       unsafe_allow_html=True)
        elif risk_level == "Trung bình":
            st.markdown(f'<div class="warning-box"><b>⚠️ {assessment}</b><br/>Mức độ rủi ro: {risk_level}</div>', 
                       unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="danger-box"><b>🚨 {assessment}</b><br/>Mức độ rủi ro: {risk_level}</div>', 
                       unsafe_allow_html=True)
        
        # Bảng chi tiết
        st.markdown("#### Chi tiết tài chính")
        
        detail_data = {
            "Chỉ tiêu": [
                "Số tiền vay",
                "Lãi suất",
                "Thời hạn",
                "Trả nợ hàng tháng",
                "Tổng lãi phải trả",
                "Tổng số tiền phải trả",
                "Thu nhập tháng",
                "Chi phí tháng",
                "Nợ khác",
                "Dòng tiền ròng",
                "DSR",
                "Biên an toàn"
            ],
            "Giá trị": [
                f"{format_number(summary['loan_amount'])} VND",
                f"{summary['interest_rate']}% /năm",
                f"{summary['loan_term']} tháng",
                f"{format_number(summary['monthly_payment'])} VND",
                f"{format_number(summary['total_interest'])} VND",
                f"{format_number(summary['total_payment'])} VND",
                f"{format_number(summary['monthly_income'])} VND",
                f"{format_number(summary['monthly_expense'])} VND",
                f"{format_number(summary['other_debt'])} VND",
                f"{format_number(summary['net_cash_flow'])} VND",
                f"{summary['dsr']:.2f}%" if summary['dsr'] >= 0 else "N/A",
                f"{summary['safety_margin']:.2f}%"
            ]
        }
        
        import pandas as pd
        df = pd.DataFrame(detail_data)
        st.dataframe(df, use_container_width=True, hide_index=True)
        
        # Lưu vào session state để dùng cho các tab khác
        st.session_state.financial_summary = summary
        st.session_state.payment_schedule = calc.calculate_payment_schedule()
    else:
        st.info("ℹ️ Vui lòng nhập thông tin khoản vay để tính toán")


def render_tab_charts():
    """Tab 5: Biểu đồ"""
    st.markdown("### 📈 Biểu Đồ Phân Tích")
    
    if 'financial_summary' not in st.session_state:
        st.info("ℹ️ Vui lòng tính toán các chỉ tiêu tài chính trước")
        return
    
    chart_gen = ChartGenerator()
    
    # Chọn loại biểu đồ
    chart_type = st.selectbox(
        "Chọn loại biểu đồ",
        [
            "Lịch trả nợ hàng tháng",
            "Phân tích dòng tiền",
            "Cơ cấu nguồn vốn",
            "So sánh thu nhập và nghĩa vụ",
            "Dư nợ giảm dần"
        ]
    )
    
    try:
        if chart_type == "Lịch trả nợ hàng tháng":
            chart_buf = chart_gen.plot_payment_schedule(st.session_state.payment_schedule)
            st.image(chart_buf, use_container_width=True)
        
        elif chart_type == "Phân tích dòng tiền":
            chart_buf = chart_gen.plot_cash_flow(
                st.session_state.payment_schedule,
                st.session_state.financial_info['monthly_income'],
                st.session_state.financial_info['monthly_expense']
            )
            st.image(chart_buf, use_container_width=True)
        
        elif chart_type == "Cơ cấu nguồn vốn":
            chart_buf = chart_gen.plot_capital_allocation(
                st.session_state.loan_info['total_need'],
                st.session_state.loan_info['equity'],
                st.session_state.loan_info['loan_amount']
            )
            st.image(chart_buf, use_container_width=True)
        
        elif chart_type == "So sánh thu nhập và nghĩa vụ":
            chart_buf = chart_gen.plot_debt_ratio(
                st.session_state.financial_info['monthly_income'],
                st.session_state.financial_summary['monthly_payment'],
                st.session_state.financial_info['monthly_expense'],
                st.session_state.financial_info['other_debt']
            )
            st.image(chart_buf, use_container_width=True)
        
        elif chart_type == "Dư nợ giảm dần":
            chart_buf = chart_gen.plot_remaining_balance(st.session_state.payment_schedule)
            st.image(chart_buf, use_container_width=True)
        
    except Exception as e:
        st.error(f"Lỗi khi vẽ biểu đồ: {str(e)}")


def render_tab_ai_analysis():
    """Tab 6: Phân tích AI"""
    st.markdown("### 🤖 Phân Tích AI - Gemini")
    
    if not st.session_state.get('api_key'):
        st.warning("⚠️ Vui lòng nhập API key ở sidebar để sử dụng tính năng này")
        return
    
    gemini_client = get_gemini_client(st.session_state.api_key)
    
    if not gemini_client:
        st.error("❌ Không thể kết nối với Gemini API")
        return
    
    # Phần 1: Phân tích từ file
    st.markdown("#### 📄 Phần 1: Phân tích từ File Upload")
    st.caption("Nguồn dữ liệu: File phương án khách hàng upload")
    
    if st.button("🔍 Phân Tích File", use_container_width=True):
        if st.session_state.raw_file_content:
            with st.spinner("Đang phân tích file..."):
                analysis = gemini_client.analyze_from_file(st.session_state.raw_file_content)
                st.session_state.file_analysis = analysis
        else:
            st.warning("⚠️ Chưa có file nào được upload")
    
    if st.session_state.file_analysis:
        st.markdown(st.session_state.file_analysis)
    
    st.markdown("---")
    
    # Phần 2: Phân tích từ dữ liệu đã nhập
    st.markdown("#### ✏️ Phần 2: Phân tích từ Dữ Liệu Đã Nhập/Chỉnh Sửa")
    st.caption("Nguồn dữ liệu: Dữ liệu sau khi hiệu chỉnh tại giao diện")
    
    if st.button("🔍 Phân Tích Dữ Liệu Hiện Tại", use_container_width=True):
        if 'financial_summary' in st.session_state:
            with st.spinner("Đang phân tích dữ liệu..."):
                data_for_analysis = {
                    'customer_name': st.session_state.customer_info['name'],
                    'customer_cccd': st.session_state.customer_info['cccd'],
                    'customer_address': st.session_state.customer_info['address'],
                    'loan_purpose': st.session_state.loan_info['purpose'],
                    'loan_amount': st.session_state.loan_info['loan_amount'],
                    'interest_rate': st.session_state.loan_info['interest_rate'],
                    'loan_term': st.session_state.loan_info['loan_term'],
                    'monthly_payment': st.session_state.financial_summary['monthly_payment'],
                    'monthly_income': st.session_state.financial_info['monthly_income'],
                    'monthly_expense': st.session_state.financial_info['monthly_expense'],
                    'net_cash_flow': st.session_state.financial_summary['net_cash_flow'],
                    'dsr': st.session_state.financial_summary['dsr'],
                    'safety_margin': st.session_state.financial_summary['safety_margin'],
                    'collateral_type': st.session_state.collateral_info['asset_type'],
                    'collateral_value': st.session_state.collateral_info['market_value'],
                    'ltv': st.session_state.financial_summary.get('ltv', 0)
                }
                
                analysis = gemini_client.analyze_from_data(data_for_analysis)
                st.session_state.data_analysis = analysis
        else:
            st.warning("⚠️ Vui lòng tính toán các chỉ tiêu tài chính trước")
    
    if st.session_state.data_analysis:
        st.markdown(st.session_state.data_analysis)


def render_tab_chatbot():
    """Tab 7: Chatbot Gemini"""
    st.markdown("### 💬 Chatbox AI - Hỏi Đáp Với Gemini")
    
    if not st.session_state.get('api_key'):
        st.warning("⚠️ Vui lòng nhập API key ở sidebar để sử dụng tính năng này")
        return
    
    gemini_client = get_gemini_client(st.session_state.api_key)
    
    if not gemini_client:
        st.error("❌ Không thể kết nối với Gemini API")
        return
    
    # Nút xóa lịch sử
    if st.button("🗑️ Xóa hội thoại", use_container_width=True):
        st.session_state.chat_history = []
        st.rerun()
    
    # Hiển thị lịch sử chat
    chat_container = st.container()
    with chat_container:
        for msg in st.session_state.chat_history:
            if msg['role'] == 'user':
                st.markdown(f"**👤 Bạn:** {msg['content']}")
            else:
                st.markdown(f"**🤖 Gemini:** {msg['content']}")
            st.markdown("---")
    
    # Input box
    user_input = st.text_input(
        "Nhập câu hỏi của bạn",
        key='chat_input',
        placeholder="Ví dụ: Phân tích rủi ro của khoản vay này?"
    )
    
    col1, col2 = st.columns([6, 1])
    with col2:
        send_button = st.button("📤 Gửi", use_container_width=True)
    
    if send_button and user_input:
        # Thêm tin nhắn người dùng
        st.session_state.chat_history.append({
            'role': 'user',
            'content': user_input
        })
        
        # Lấy phản hồi từ Gemini
        with st.spinner("Đang suy nghĩ..."):
            response = gemini_client.chat(user_input, st.session_state.chat_history)
        
        # Thêm phản hồi
        st.session_state.chat_history.append({
            'role': 'assistant',
            'content': response
        })
        
        st.rerun()


def render_tab_export():
    """Tab 8: Xuất file"""
    st.markdown("### 📥 Xuất Dữ Liệu")
    
    if 'financial_summary' not in st.session_state:
        st.info("ℹ️ Vui lòng tính toán các chỉ tiêu tài chính trước")
        return
    
    export_type = st.selectbox(
        "Chọn loại xuất dữ liệu",
        [
            "Xuất bảng kê kế hoạch trả nợ (Excel)",
            "Xuất báo cáo thẩm định (PDF)"
        ]
    )
    
    if export_type == "Xuất bảng kê kế hoạch trả nợ (Excel)":
        st.markdown("#### 📊 Bảng Kê Kế Hoạch Trả Nợ")
        
        if st.button("📥 Tạo File Excel", use_container_width=True):
            try:
                exporter = ExcelExporter()
                
                loan_info = {
                    'customer_name': st.session_state.customer_info['name'],
                    'loan_amount': st.session_state.loan_info['loan_amount'],
                    'interest_rate': st.session_state.loan_info['interest_rate'],
                    'loan_term': st.session_state.loan_info['loan_term']
                }
                
                excel_file = exporter.create_payment_schedule_excel(
                    st.session_state.payment_schedule,
                    loan_info
                )
                
                st.download_button(
                    label="⬇️ Tải xuống file Excel",
                    data=excel_file,
                    file_name=f"Ke_hoach_tra_no_{st.session_state.customer_info['name'].replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success("✅ File Excel đã được tạo thành công!")
            except Exception as e:
                st.error(f"❌ Lỗi khi tạo file: {str(e)}")
    
    elif export_type == "Xuất báo cáo thẩm định (PDF)":
        st.markdown("#### 📄 Báo Cáo Thẩm Định")
        
        if st.button("📥 Tạo Báo Cáo PDF", use_container_width=True):
            try:
                with st.spinner("Đang tạo báo cáo..."):
                    pdf_exporter = PDFExporter()
                    
                    # Chuẩn bị dữ liệu
                    report_data = {
                        'customer_name': st.session_state.customer_info['name'],
                        'customer_cccd': st.session_state.customer_info['cccd'],
                        'customer_address': st.session_state.customer_info['address'],
                        'customer_phone': st.session_state.customer_info['phone'],
                        'loan_purpose': st.session_state.loan_info['purpose'],
                        'total_need': st.session_state.loan_info['total_need'],
                        'equity': st.session_state.loan_info['equity'],
                        'loan_amount': st.session_state.loan_info['loan_amount'],
                        'interest_rate': st.session_state.loan_info['interest_rate'],
                        'loan_term': st.session_state.loan_info['loan_term'],
                        'monthly_payment': st.session_state.financial_summary['monthly_payment'],
                        'collateral_type': st.session_state.collateral_info['asset_type'],
                        'collateral_value': st.session_state.collateral_info['market_value'],
                        'asset_address': st.session_state.collateral_info['asset_address'],
                        'ltv': st.session_state.financial_summary.get('ltv', 0),
                        'legal_docs': st.session_state.collateral_info['legal_docs'],
                        'monthly_income': st.session_state.financial_info['monthly_income'],
                        'monthly_expense': st.session_state.financial_info['monthly_expense'],
                        'other_debt': st.session_state.financial_info['other_debt'],
                        'net_cash_flow': st.session_state.financial_summary['net_cash_flow'],
                        'dsr': st.session_state.financial_summary['dsr'],
                        'safety_margin': st.session_state.financial_summary['safety_margin'],
                        'assessment': st.session_state.financial_summary['assessment'],
                        'risk_level': st.session_state.financial_summary['risk_level'],
                        'can_repay': st.session_state.financial_summary['can_repay'],
                        'ai_analysis': st.session_state.get('data_analysis', '')
                    }
                    
                    pdf_file = pdf_exporter.create_assessment_report(
                        report_data,
                        st.session_state.payment_schedule
                    )
                    
                    st.download_button(
                        label="⬇️ Tải xuống Báo Cáo PDF",
                        data=pdf_file,
                        file_name=f"Bao_cao_tham_dinh_{st.session_state.customer_info['name'].replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                    st.success("✅ Báo cáo PDF đã được tạo thành công!")
            except Exception as e:
                st.error(f"❌ Lỗi khi tạo báo cáo: {str(e)}")


def main():
    """Hàm main"""
    # Khởi tạo session state
    initialize_session_state()
    
    # Header
    st.markdown('<div class="main-header">🏦 HỆ THỐNG THẨM ĐỊNH PHƯƠNG ÁN KINH DOANH</div>', 
                unsafe_allow_html=True)
    
    # Sidebar
    render_sidebar()
    
    # Hiển thị trạng thái
    if st.session_state.data_loaded:
        if st.session_state.data_modified:
            st.info("ℹ️ Dữ liệu đã được thay đổi. Các chỉ tiêu sẽ được tính toán lại.")
        else:
            st.success("✅ Dữ liệu đã được tải từ file")
    
    # Tabs
    tabs = st.tabs([
        "👤 Khách hàng",
        "💰 Khoản vay",
        "🏠 TSBĐ",
        "📊 Chỉ tiêu",
        "📈 Biểu đồ",
        "🤖 AI Phân tích",
        "💬 Chatbot",
        "📥 Xuất file"
    ])
    
    with tabs[0]:
        render_tab_customer_info()
    
    with tabs[1]:
        render_tab_loan_info()
    
    with tabs[2]:
        render_tab_collateral_info()
    
    with tabs[3]:
        render_tab_financial_calculations()
    
    with tabs[4]:
        render_tab_charts()
    
    with tabs[5]:
        render_tab_ai_analysis()
    
    with tabs[6]:
        render_tab_chatbot()
    
    with tabs[7]:
        render_tab_export()


if __name__ == "__main__":
    main()
