import pandas as pd
import numpy as np
from openpyxl import Workbook

def calculate_metrics(strategy, data):
    # Base calculations for all strategies
    purchase_price = data['Purchase Price']
    deposit = data['Deposit']
    ltv = data['LTV'] / 100
    mortgage_interest_rate = data['Mortgage Interest Rate'] / 100
    rent_income = data['Rent Income']
    expenses = data['Expenses']
    refurbishment_cost = data.get('Refurbishment Cost', 0)
    after_refurbishment_value = data.get('After Refurbishment Value', purchase_price)

    # Mortgage monthly payment (assuming interest-only mortgage)
    loan_amount = purchase_price * ltv
    monthly_mortgage_payment = loan_amount * mortgage_interest_rate / 12

    # Monthly cash flow
    monthly_cash_flow = rent_income - (monthly_mortgage_payment + expenses)

    # ROI calculation
    total_investment = deposit + refurbishment_cost
    annual_cash_flow = monthly_cash_flow * 12
    roi = (annual_cash_flow / total_investment) * 100

    # Yields
    gross_yield = (rent_income * 12 / purchase_price) * 100
    net_yield = (annual_cash_flow / purchase_price) * 100

    # BRR-specific metrics
    equity_gain = after_refurbishment_value - (loan_amount + refurbishment_cost)

    # Return results based on the strategy
    metrics = {
        'Monthly Cash Flow': monthly_cash_flow,
        'ROI (%)': roi,
        'Gross Yield (%)': gross_yield,
        'Net Yield (%)': net_yield,
        'Equity Gain': equity_gain if 'BRR' in strategy else np.nan,
    }

    return metrics

def create_excel_analyzer():
    # Define strategies and input fields
    strategies = [
        "Turnkey BTL",
        "Turnkey HMO",
        "Turnkey SA",
        "BRR BTL",
        "BRR HMO",
        "BRR SA",
    ]

    # Sample data for input fields
    input_data = {
        'Purchase Price': 100000,
        'Deposit': 25000,
        'LTV': 75,
        'Mortgage Interest Rate': 5,
        'Rent Income': 1000,
        'Expenses': 300,
        'Refurbishment Cost': 20000,
        'After Refurbishment Value': 150000,
    }

    # Create a workbook
    with pd.ExcelWriter('Property_Strategy_Analyzer.xlsx', engine='openpyxl') as writer:
        # Iterate over each strategy and calculate metrics
        for strategy in strategies:
            metrics = calculate_metrics(strategy, input_data)

            # Create a DataFrame for display
            df = pd.DataFrame({
                'Metric': metrics.keys(),
                'Value': metrics.values()
            })

            # Write to a sheet named after the strategy
            df.to_excel(writer, sheet_name=strategy, index=False)

    print("Property Strategy Analyzer Excel file has been created!")

if __name__ == "__main__":
    create_excel_analyzer()
