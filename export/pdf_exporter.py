# export/pdf_exporter.py
"""Module xuất báo cáo PDF"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.platypus import Image as RLImage
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
import io
from typing import Dict, List
from src.utils import format_number
import matplotlib.pyplot as plt
from datetime import datetime


class PDFExporter:
    """Class xuất báo cáo PDF"""
    
    def __init__(self):
        """Khởi tạo PDF exporter"""
        self.styles = getSampleStyleSheet()
        
        # Custom styles
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#1f4788'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )
        
        self.heading_style = ParagraphStyle(
            'CustomHeading',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=colors.HexColor('#1f4788'),
            spaceAfter=12,
            spaceBefore=12,
            fontName='Helvetica-Bold'
        )
        
        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontSize=11,
            spaceAfter=6,
            alignment=TA_JUSTIFY
        )
    
    def create_assessment_report(self, data: Dict, 
                                 schedule: List[Dict] = None,
                                 chart_images: Dict = None) -> io.BytesIO:
        """
        Tạo báo cáo thẩm định PDF
        
        Args:
            data: Dữ liệu phương án
            schedule: Lịch trả nợ (optional)
            chart_images: Dictionary chứa các biểu đồ dạng BytesIO (optional)
            
        Returns:
            BytesIO object chứa file PDF
        """
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )
        
        # Container cho các elements
        story = []
        
        # Tiêu đề
        title = Paragraph("BÁO CÁO THẨM ĐỊNH PHƯƠNG ÁN KINH DOANH", self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.5*cm))
        
        # Ngày báo cáo
        date_text = f"Ngày báo cáo: {datetime.now().strftime('%d/%m/%Y')}"
        story.append(Paragraph(date_text, self.normal_style))
        story.append(Spacer(1, 0.5*cm))
        
        # I. THÔNG TIN KHÁCH HÀNG
        story.append(Paragraph("I. THÔNG TIN KHÁCH HÀNG", self.heading_style))
        
        customer_data = [
            ['Họ và tên:', data.get('customer_name', 'N/A')],
            ['CCCD/CMND:', data.get('customer_cccd', 'N/A')],
            ['Địa chỉ:', data.get('customer_address', 'N/A')],
            ['Số điện thoại:', data.get('customer_phone', 'N/A')],
        ]
        
        customer_table = Table(customer_data, colWidths=[5*cm, 12*cm])
        customer_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(customer_table)
        story.append(Spacer(1, 0.5*cm))
        
        # II. THÔNG TIN KHOẢN VAY
        story.append(Paragraph("II. THÔNG TIN KHOẢN VAY", self.heading_style))
        
        loan_data = [
            ['Mục đích vay:', data.get('loan_purpose', 'N/A')],
            ['Tổng nhu cầu vốn:', f"{format_number(data.get('total_need', 0))} VND"],
            ['Vốn đối ứng:', f"{format_number(data.get('equity', 0))} VND"],
            ['Số tiền vay:', f"{format_number(data.get('loan_amount', 0))} VND"],
            ['Lãi suất:', f"{data.get('interest_rate', 0)}% /năm"],
            ['Thời hạn vay:', f"{data.get('loan_term', 0)} tháng"],
            ['Trả nợ hàng tháng:', f"{format_number(data.get('monthly_payment', 0))} VND"],
        ]
        
        loan_table = Table(loan_data, colWidths=[5*cm, 12*cm])
        loan_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(loan_table)
        story.append(Spacer(1, 0.5*cm))
        
        # III. TÀI SẢN BẢO ĐẢM
        story.append(Paragraph("III. TÀI SẢN BẢO ĐẢM", self.heading_style))
        
        collateral_data = [
            ['Loại tài sản:', data.get('collateral_type', 'N/A')],
            ['Giá trị thị trường:', f"{format_number(data.get('collateral_value', 0))} VND"],
            ['Địa chỉ tài sản:', data.get('asset_address', 'N/A')],
            ['LTV:', f"{data.get('ltv', 0):.2f}%"],
            ['Giấy tờ pháp lý:', data.get('legal_docs', 'N/A')],
        ]
        
        collateral_table = Table(collateral_data, colWidths=[5*cm, 12*cm])
        collateral_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(collateral_table)
        story.append(Spacer(1, 0.5*cm))
        
        # IV. CHỈ TIÊU TÀI CHÍNH
        story.append(Paragraph("IV. CHỈ TIÊU TÀI CHÍNH", self.heading_style))
        
        financial_data = [
            ['Thu nhập hàng tháng:', f"{format_number(data.get('monthly_income', 0))} VND"],
            ['Chi phí hàng tháng:', f"{format_number(data.get('monthly_expense', 0))} VND"],
            ['Nghĩa vụ nợ khác:', f"{format_number(data.get('other_debt', 0))} VND"],
            ['Dòng tiền ròng:', f"{format_number(data.get('net_cash_flow', 0))} VND"],
            ['Tỷ lệ DSR:', f"{data.get('dsr', 0):.2f}%"],
            ['Biên an toàn:', f"{data.get('safety_margin', 0):.2f}%"],
        ]
        
        financial_table = Table(financial_data, colWidths=[5*cm, 12*cm])
        financial_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(financial_table)
        story.append(Spacer(1, 0.5*cm))
        
        # V. ĐÁNH GIÁ
        story.append(Paragraph("V. ĐÁNH GIÁ NĂNG LỰC TRẢ NỢ", self.heading_style))
        
        assessment_text = f"""
        <b>Kết luận:</b> {data.get('assessment', 'N/A')}<br/>
        <b>Mức độ rủi ro:</b> {data.get('risk_level', 'N/A')}<br/>
        <b>Khả năng trả nợ:</b> {'Đủ khả năng' if data.get('can_repay', False) else 'Không đủ khả năng'}
        """
        story.append(Paragraph(assessment_text, self.normal_style))
        story.append(Spacer(1, 0.5*cm))
        
        # VI. PHÂN TÍCH AI (nếu có)
        if data.get('ai_analysis'):
            story.append(PageBreak())
            story.append(Paragraph("VI. PHÂN TÍCH CHUYÊN SÂU (AI)", self.heading_style))
            
            analysis_paragraphs = data['ai_analysis'].split('\n')
            for para in analysis_paragraphs:
                if para.strip():
                    story.append(Paragraph(para.strip(), self.normal_style))
                    story.append(Spacer(1, 0.2*cm))
        
        # VII. BIỂU ĐỒ (nếu có)
        if chart_images:
            story.append(PageBreak())
            story.append(Paragraph("VII. BIỂU ĐỒ PHÂN TÍCH", self.heading_style))
            
            for chart_name, chart_buffer in chart_images.items():
                if chart_buffer:
                    chart_buffer.seek(0)
                    img = RLImage(chart_buffer, width=15*cm, height=10*cm)
                    story.append(img)
                    story.append(Spacer(1, 0.5*cm))
        
        # Build PDF
        doc.build(story)
        buffer.seek(0)
        return buffer
