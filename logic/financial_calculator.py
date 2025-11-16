# logic/financial_calculator.py
"""Module tính toán các chỉ tiêu tài chính"""

import math
from typing import Dict, List, Tuple
from src.utils import safe_divide


class FinancialCalculator:
    """Class tính toán các chỉ tiêu tài chính"""
    
    def __init__(self, loan_amount: float, interest_rate: float, 
                 loan_term: int, monthly_income: float = 0,
                 monthly_expense: float = 0, other_debt: float = 0):
        """
        Khởi tạo calculator
        
        Args:
            loan_amount: Số tiền vay (VND)
            interest_rate: Lãi suất năm (%)
            loan_term: Thời hạn vay (tháng)
            monthly_income: Thu nhập hàng tháng
            monthly_expense: Chi phí hàng tháng
            other_debt: Nợ khác hàng tháng
        """
        self.loan_amount = loan_amount
        self.interest_rate = interest_rate
        self.loan_term = loan_term
        self.monthly_income = monthly_income
        self.monthly_expense = monthly_expense
        self.other_debt = other_debt
        
        # Lãi suất tháng
        self.monthly_rate = interest_rate / 100 / 12
    
    def calculate_monthly_payment(self) -> float:
        """
        Tính toán trả nợ hàng tháng (gốc + lãi) theo phương thức dư nợ giảm dần
        
        Returns:
            Số tiền trả nợ hàng tháng
        """
        if self.loan_amount <= 0 or self.loan_term <= 0:
            return 0.0
        
        # Trả gốc đều hàng tháng
        principal_payment = self.loan_amount / self.loan_term
        
        # Lãi tháng đầu (cao nhất)
        interest_payment = self.loan_amount * self.monthly_rate
        
        # Tổng trả tháng đầu
        return principal_payment + interest_payment
    
    def calculate_payment_schedule(self) -> List[Dict[str, float]]:
        """
        Tính toán lịch trả nợ chi tiết theo từng tháng
        Phương thức: Dư nợ giảm dần (trả gốc đều)
        
        Returns:
            Danh sách các kỳ trả nợ
        """
        schedule = []
        remaining_balance = self.loan_amount
        monthly_principal = self.loan_amount / self.loan_term
        
        for month in range(1, self.loan_term + 1):
            # Tính lãi trên dư nợ đầu kỳ
            interest = remaining_balance * self.monthly_rate
            
            # Tính gốc (cố định)
            principal = monthly_principal
            
            # Tổng trả = gốc + lãi
            total_payment = principal + interest
            
            # Dư nợ sau khi trả
            remaining_balance -= principal
            
            # Đảm bảo dư nợ không âm do làm tròn
            if remaining_balance < 1:
                remaining_balance = 0
            
            schedule.append({
                'month': month,
                'principal': principal,
                'interest': interest,
                'total_payment': total_payment,
                'remaining_balance': remaining_balance
            })
        
        return schedule
    
    def calculate_total_interest(self) -> float:
        """Tính tổng lãi phải trả"""
        schedule = self.calculate_payment_schedule()
        return sum(period['interest'] for period in schedule)
    
    def calculate_total_payment(self) -> float:
        """Tính tổng số tiền phải trả (gốc + lãi)"""
        return self.loan_amount + self.calculate_total_interest()
    
    def calculate_dsr(self) -> float:
        """
        Tính Debt Service Ratio (DSR) - Tỷ lệ trả nợ trên thu nhập
        
        Returns:
            DSR (%), giá trị < 0 nếu không tính được
        """
        if self.monthly_income <= 0:
            return -1.0
        
        monthly_payment = self.calculate_monthly_payment()
        total_debt = monthly_payment + self.other_debt
        
        dsr = (total_debt / self.monthly_income) * 100
        return dsr
    
    def calculate_net_cash_flow(self) -> float:
        """
        Tính dòng tiền ròng hàng tháng
        
        Returns:
            Dòng tiền ròng (VND)
        """
        monthly_payment = self.calculate_monthly_payment()
        net_flow = self.monthly_income - self.monthly_expense - monthly_payment - self.other_debt
        return net_flow
    
    def calculate_safety_margin(self) -> float:
        """
        Tính biên độ an toàn tài chính
        
        Returns:
            Tỷ lệ % dòng tiền ròng trên thu nhập
        """
        if self.monthly_income <= 0:
            return 0.0
        
        net_flow = self.calculate_net_cash_flow()
        margin = (net_flow / self.monthly_income) * 100
        return margin
    
    def assess_repayment_capacity(self) -> Dict[str, any]:
        """
        Đánh giá năng lực trả nợ
        
        Returns:
            Dictionary chứa kết quả đánh giá
        """
        dsr = self.calculate_dsr()
        net_flow = self.calculate_net_cash_flow()
        margin = self.calculate_safety_margin()
        
        # Đánh giá
        if dsr < 0:
            assessment = "Không đủ dữ liệu"
            risk_level = "N/A"
        elif dsr <= 40:
            assessment = "Tốt - Khả năng trả nợ cao"
            risk_level = "Thấp"
        elif dsr <= 60:
            assessment = "Trung bình - Cần theo dõi"
            risk_level = "Trung bình"
        else:
            assessment = "Yếu - Rủi ro cao"
            risk_level = "Cao"
        
        return {
            'dsr': dsr,
            'net_cash_flow': net_flow,
            'safety_margin': margin,
            'assessment': assessment,
            'risk_level': risk_level,
            'can_repay': net_flow > 0 and (dsr <= 60 if dsr > 0 else False)
        }
    
    def calculate_ltv(self, collateral_value: float) -> float:
        """
        Tính Loan to Value ratio
        
        Args:
            collateral_value: Giá trị tài sản bảo đảm
            
        Returns:
            LTV (%)
        """
        if collateral_value <= 0:
            return 0.0
        
        ltv = (self.loan_amount / collateral_value) * 100
        return ltv
    
    def get_summary(self, collateral_value: float = 0) -> Dict[str, any]:
        """
        Lấy tóm tắt tất cả các chỉ tiêu
        
        Args:
            collateral_value: Giá trị tài sản bảo đảm
            
        Returns:
            Dictionary chứa tất cả chỉ tiêu
        """
        monthly_payment = self.calculate_monthly_payment()
        total_interest = self.calculate_total_interest()
        total_payment = self.calculate_total_payment()
        capacity = self.assess_repayment_capacity()
        
        summary = {
            'loan_amount': self.loan_amount,
            'interest_rate': self.interest_rate,
            'loan_term': self.loan_term,
            'monthly_payment': monthly_payment,
            'total_interest': total_interest,
            'total_payment': total_payment,
            'monthly_income': self.monthly_income,
            'monthly_expense': self.monthly_expense,
            'other_debt': self.other_debt,
            **capacity
        }
        
        if collateral_value > 0:
            summary['ltv'] = self.calculate_ltv(collateral_value)
            summary['collateral_value'] = collateral_value
        
        return summary
