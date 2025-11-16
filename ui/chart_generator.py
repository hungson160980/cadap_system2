# ui/chart_generator.py
"""Module tạo biểu đồ"""

import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Sử dụng backend không cần GUI
import io
from typing import List, Dict
from src.utils import format_number

# Thiết lập font tiếng Việt
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False


class ChartGenerator:
    """Class tạo các biểu đồ phân tích"""
    
    def __init__(self):
        """Khởi tạo chart generator"""
        self.fig_size = (10, 6)
        self.dpi = 100
    
    def plot_payment_schedule(self, schedule: List[Dict]) -> io.BytesIO:
        """
        Vẽ biểu đồ lịch trả nợ
        
        Args:
            schedule: Lịch trả nợ
            
        Returns:
            BytesIO chứa hình ảnh
        """
        months = [p['month'] for p in schedule]
        principals = [p['principal'] for p in schedule]
        interests = [p['interest'] for p in schedule]
        
        fig, ax = plt.subplots(figsize=self.fig_size, dpi=self.dpi)
        
        ax.plot(months, principals, label='Gốc', linewidth=2, marker='o', markersize=3)
        ax.plot(months, interests, label='Lãi', linewidth=2, marker='s', markersize=3)
        
        ax.set_xlabel('Tháng', fontsize=12)
        ax.set_ylabel('Số tiền (VND)', fontsize=12)
        ax.set_title('Lịch Trả Nợ Hàng Tháng', fontsize=14, fontweight='bold')
        ax.legend(fontsize=10)
        ax.grid(True, alpha=0.3)
        
        # Format trục y
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_number(x)))
        
        plt.tight_layout()
        
        # Lưu vào BytesIO
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)
        
        return buf
    
    def plot_cash_flow(self, schedule: List[Dict], 
                       monthly_income: float,
                       monthly_expense: float) -> io.BytesIO:
        """
        Vẽ biểu đồ dòng tiền
        
        Args:
            schedule: Lịch trả nợ
            monthly_income: Thu nhập tháng
            monthly_expense: Chi phí tháng
            
        Returns:
            BytesIO chứa hình ảnh
        """
        months = [p['month'] for p in schedule]
        payments = [p['total_payment'] for p in schedule]
        net_flows = [monthly_income - monthly_expense - p for p in payments]
        
        fig, ax = plt.subplots(figsize=self.fig_size, dpi=self.dpi)
        
        # Vẽ thu nhập
        ax.axhline(y=monthly_income, color='green', linestyle='--', 
                   label='Thu nhập', linewidth=2)
        
        # Vẽ chi phí + trả nợ
        total_outflow = [monthly_expense + p for p in payments]
        ax.plot(months, total_outflow, label='Chi phí + Trả nợ', 
                linewidth=2, color='red', marker='o', markersize=3)
        
        # Vẽ dòng tiền ròng
        ax.plot(months, net_flows, label='Dòng tiền ròng', 
                linewidth=2, color='blue', marker='s', markersize=3)
        
        # Vẽ vùng dương/âm
        ax.fill_between(months, 0, net_flows, where=[nf >= 0 for nf in net_flows],
                        alpha=0.3, color='green', label='Dương')
        ax.fill_between(months, 0, net_flows, where=[nf < 0 for nf in net_flows],
                        alpha=0.3, color='red', label='Âm')
        
        ax.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
        
        ax.set_xlabel('Tháng', fontsize=12)
        ax.set_ylabel('Số tiền (VND)', fontsize=12)
        ax.set_title('Phân Tích Dòng Tiền', fontsize=14, fontweight='bold')
        ax.legend(fontsize=9)
        ax.grid(True, alpha=0.3)
        
        # Format trục y
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_number(x)))
        
        plt.tight_layout()
        
        # Lưu vào BytesIO
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)
        
        return buf
    
    def plot_capital_allocation(self, total_need: float, 
                                equity: float, 
                                loan_amount: float) -> io.BytesIO:
        """
        Vẽ biểu đồ phân bổ vốn
        
        Args:
            total_need: Tổng nhu cầu vốn
            equity: Vốn đối ứng
            loan_amount: Số tiền vay
            
        Returns:
            BytesIO chứa hình ảnh
        """
        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        
        sizes = [equity, loan_amount]
        labels = [
            f'Vốn đối ứng\n{format_number(equity)} VND\n({equity/total_need*100:.1f}%)',
            f'Vốn vay\n{format_number(loan_amount)} VND\n({loan_amount/total_need*100:.1f}%)'
        ]
        colors = ['#2ecc71', '#3498db']
        explode = (0.05, 0.05)
        
        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
               startangle=90, explode=explode, textprops={'fontsize': 11})
        
        ax.set_title('Cơ Cấu Nguồn Vốn', fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        
        # Lưu vào BytesIO
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)
        
        return buf
    
    def plot_debt_ratio(self, monthly_income: float,
                       monthly_payment: float,
                       monthly_expense: float,
                       other_debt: float) -> io.BytesIO:
        """
        Vẽ biểu đồ tỷ lệ thu nhập / nghĩa vụ
        
        Args:
            monthly_income: Thu nhập tháng
            monthly_payment: Trả nợ tháng
            monthly_expense: Chi phí tháng
            other_debt: Nợ khác
            
        Returns:
            BytesIO chứa hình ảnh
        """
        fig, ax = plt.subplots(figsize=self.fig_size, dpi=self.dpi)
        
        categories = ['Thu nhập', 'Chi phí', 'Trả nợ vay', 'Nợ khác', 'Còn lại']
        values = [
            monthly_income,
            monthly_expense,
            monthly_payment,
            other_debt,
            max(0, monthly_income - monthly_expense - monthly_payment - other_debt)
        ]
        colors_list = ['#2ecc71', '#e74c3c', '#3498db', '#f39c12', '#95a5a6']
        
        bars = ax.bar(categories, values, color=colors_list, edgecolor='black', linewidth=1.5)
        
        # Thêm giá trị trên mỗi cột
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{format_number(height)}',
                   ha='center', va='bottom', fontsize=10)
        
        ax.set_ylabel('Số tiền (VND)', fontsize=12)
        ax.set_title('So Sánh Thu Nhập và Nghĩa Vụ Tài Chính', 
                    fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3, axis='y')
        
        # Format trục y
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_number(x)))
        
        # Xoay nhãn trục x
        plt.xticks(rotation=15, ha='right')
        
        plt.tight_layout()
        
        # Lưu vào BytesIO
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)
        
        return buf
    
    def plot_remaining_balance(self, schedule: List[Dict]) -> io.BytesIO:
        """
        Vẽ biểu đồ dư nợ giảm dần
        
        Args:
            schedule: Lịch trả nợ
            
        Returns:
            BytesIO chứa hình ảnh
        """
        months = [p['month'] for p in schedule]
        balances = [p['remaining_balance'] for p in schedule]
        
        # Thêm điểm đầu (dư nợ ban đầu)
        if schedule:
            months = [0] + months
            initial_balance = schedule[0]['remaining_balance'] + schedule[0]['principal']
            balances = [initial_balance] + balances
        
        fig, ax = plt.subplots(figsize=self.fig_size, dpi=self.dpi)
        
        ax.plot(months, balances, linewidth=2.5, color='#e74c3c', marker='o', markersize=4)
        ax.fill_between(months, 0, balances, alpha=0.3, color='#e74c3c')
        
        ax.set_xlabel('Tháng', fontsize=12)
        ax.set_ylabel('Dư nợ (VND)', fontsize=12)
        ax.set_title('Dư Nợ Giảm Dần Theo Thời Gian', fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        
        # Format trục y
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_number(x)))
        
        plt.tight_layout()
        
        # Lưu vào BytesIO
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)
        
        return buf
