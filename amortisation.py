import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
from openpyxl import Workbook

# Function to create input template
def create_input_template():
    data = {
        'Sr no': [1],
        'bank Name or loan no': ['A1'],
        'Loan Amount': [15],
        'interest Rate': [12],
        'Loan term': [36],
        'Start Date': ['22/02/2022'],
        'payment Frequecy': ['Monthly'],
        'payment Amount': [0]
    }
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    output.seek(0)
    return output

# Function to determine financial year for a given date
def get_financial_year(date):
    if date.month >= 4:
        return date.year
    return date.year - 1

# Function to format financial year (e.g., FY 2024-25)
def format_fy(year):
    return f"FY {year}-{str(year + 1)[-2:]}"

# Function to parse date from string (dd/mm/yyyy or mm/dd/yyyy)
def parse_date(date_str):
    try:
        return datetime.strptime(date_str, '%m/%d/%Y')
    except ValueError:
        return datetime.strptime(date_str, '%d/%m/%Y')

# Function to calculate amortization schedule
def calculate_amortization(loan_amount_lakhs, annual_rate, term_months, start_date, payment_amount_lakhs=None):
    if loan_amount_lakhs <= 0 or annual_rate <= 0 or annual_rate >= 50 or term_months <= 0:
        raise ValueError("Invalid input: Loan Amount, Interest Rate, and Loan Term must be positive; Interest Rate must be < 50%.")
    
    loan_amount = loan_amount_lakhs * 100000
    monthly_rate = annual_rate / 100 / 12
    n = term_months
    
    # Calculate EMI if not provided
    if payment_amount_lakhs is None or payment_amount_lakhs == 0:
        emi = loan_amount * (monthly_rate * (1 + monthly_rate) ** n) / ((1 + monthly_rate) ** n - 1)
    else:
        emi = payment_amount_lakhs * 100000
    
    schedule = []
    balance = loan_amount
    current_date = datetime.strptime(start_date, '%Y-%m-%d')
    
    for period in range(n):
        interest = balance * monthly_rate
        principal = emi - interest
        if balance < principal:
            principal = balance
            emi = principal + interest
        balance -= principal
        
        financial_year = get_financial_year(current_date)
        
        schedule.append({
            'Date': current_date.strftime('%d/%m/%Y'),
            'Financial_Year': financial_year,
            'Principal_Lakhs': principal / 100000,
            'Interest_Lakhs': interest / 100000,
            'Payment_Lakhs': emi / 100000,
            'Balance_Lakhs': balance / 100000,
            'EMI_Lakhs': emi / 100000
        })
        
        current_date += relativedelta(months=1)
    
    return pd.DataFrame(schedule)

# Function to calculate annual metrics and pivot for output
def calculate_annual_metrics(schedule_df, sr_no, loan_name, loan_amount_lakhs):
    # Convert schedule dates to datetime for comparison
    schedule_df['Date'] = schedule_df['Date'].apply(parse_date)
    
    # Determine financial year range (cap at 20 years from earliest year)
    min_year = get_financial_year(schedule_df['Date'].min())
    max_year = min(max(schedule_df['Financial_Year']) + 1, min_year + 20)
    years = list(range(min_year, max_year))
    
    # Group by financial year for principal and interest
    annual_summary = schedule_df.groupby('Financial_Year').agg({
        'Principal_Lakhs': 'sum',
        'Interest_Lakhs': 'sum'
    }).reset_index()
    
    # Calculate outstanding balance (balance on March 31 of each year)
    outstanding_balances = []
    start_date = schedule_df['Date'].iloc[0]
    for year in years:
        end_date = datetime(year + 1, 3, 31)
        if start_date > end_date:
            outstanding_balances.append(loan_amount_lakhs)
        else:
            last_payment = schedule_df[schedule_df['Date'] <= end_date].tail(1)
            outstanding_balances.append(last_payment['Balance_Lakhs'].iloc[0] if not last_payment.empty else 0.0)
    
    # Calculate current liability (principal payments for next 12 months from March 31)
    current_liabilities = []
    for year in years:
        current_date = datetime(year + 1, 3, 31)
        next_year_end = current_date + relativedelta(years=1)
        next_12_months = schedule_df[
            (schedule_df['Date'] > current_date) & 
            (schedule_df['Date'] <= next_year_end)
        ]
        current_liabilities.append(next_12_months['Principal_Lakhs'].sum())
    
    # Create pivot DataFrames with consistent year range
    principal_pivot = pd.DataFrame({
        'Sr no': [sr_no],
        'Loan name': [loan_name]
    })
    interest_pivot = pd.DataFrame({
        'Sr no': [sr_no],
        'Loan name': [loan_name]
    })
    outstanding_pivot = pd.DataFrame({
        'Sr no': [sr_no],
        'Loan name': [loan_name]
    })
    liabilities_pivot = pd.DataFrame({
        'Sr no': [sr_no],
        'Loan name': [loan_name]
    })
    
    # Add financial year columns
    for year in years:
        fy_label = format_fy(year)
        # Principal
        principal_value = annual_summary[annual_summary['Financial_Year'] == year]['Principal_Lakhs'].iloc[0] if year in annual_summary['Financial_Year'].values else 0.0
        principal_pivot[f"Principal {fy_label}"] = [principal_value]
        
        # Interest
        interest_value = annual_summary[annual_summary['Financial_Year'] == year]['Interest_Lakhs'].iloc[0] if year in annual_summary['Financial_Year'].values else 0.0
        interest_pivot[f"Interest {fy_label}"] = [interest_value]
        
        # Outstanding
        outstanding_pivot[f"Outstanding {fy_label}"] = [outstanding_balances[years.index(year)]]
        
        # Liabilities
        liabilities_pivot[f"Liability {fy_label}"] = [current_liabilities[years.index(year)]]
    
    return schedule_df, principal_pivot, interest_pivot, outstanding_pivot, liabilities_pivot

# Function to create Excel file with one sheet for all loans and totals
def create_excel_file(data_frames, file_name, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_df = pd.concat(data_frames, ignore_index=True)
        
        # Calculate row-wise totals (sum across financial years for each loan)
        fy_columns = [col for col in combined_df.columns if col.startswith(sheet_name)]
        combined_df['Total'] = combined_df[fy_columns].sum(axis=1)
        
        # Calculate column-wise totals (sum across loans for each financial year)
        total_row = {'Sr no': 'Total', 'Loan name': ''}
        for col in fy_columns:
            total_row[col] = combined_df[col].sum()
        total_row['Total'] = combined_df['Total'].sum()
        total_df = pd.DataFrame([total_row])
        
        # Combine data and totals
        final_df = pd.concat([combined_df, total_df], ignore_index=True)
        final_df.round(2).to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    return output

# Function to create amortization schedule Excel file (one sheet per loan)
def create_schedule_excel_file(data_frames, file_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for idx, df in enumerate(data_frames, 1):
            df.round(2).to_excel(writer, sheet_name=f'Schedule_{idx}', index=False)
    output.seek(0)
    return output

# Streamlit app
st.title("Amortization Schedule Calculator (Financial Year: Apr 1 - Mar 31)")

# Download input template
st.download_button(
    label="Download Input Template",
    data=create_input_template(),
    file_name="input_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Input method selection
input_method = st.radio("Choose input method:", ["Enter Single Loan Details", "Upload Excel File"])

results = []  # Store (schedule_df, principal_pivot, interest_pivot, outstanding_pivot, liabilities_pivot) for each loan

if input_method == "Enter Single Loan Details":
    st.subheader("Enter Loan Details (Amounts in Lakhs)")
    with st.form(key='loan_form'):
        sr_no = st.text_input("Sr no", value="1")
        loan_name = st.text_input("Bank Name or Loan No", value="A1")
        loan_amount_lakhs = st.number_input("Loan Amount (Lakhs)", min_value=0.01, value=15.0, step=0.1)
        interest_rate = st.number_input("Interest Rate (%)", min_value=0.01, max_value=50.0, value=12.0, step=0.1)
        loan_term_months = st.number_input("Loan Term (Months)", min_value=1, value=36, step=1)
        start_date = st.date_input("Start Date", value=datetime(2022, 2, 22))
        payment_amount_lakhs = st.number_input("Payment Amount (Lakhs, Optional, leave 0 to calculate)", min_value=0.0, value=0.0, step=0.01)
        submit_button = st.form_submit_button(label="Calculate")
    
    if submit_button:
        try:
            start_date_str = start_date.strftime('%Y-%m-%d')
            schedule_df, principal_pivot, interest_pivot, outstanding_pivot, liabilities_pivot = calculate_annual_metrics(
                calculate_amortization(loan_amount_lakhs, interest_rate, loan_term_months, start_date_str, payment_amount_lakhs),
                sr_no, loan_name, loan_amount_lakhs
            )
            results.append((schedule_df, principal_pivot, interest_pivot, outstanding_pivot, liabilities_pivot))
            
            # Generate and provide download buttons
            st.download_button(
                label="Download Annual Principal Repayment",
                data=create_excel_file([res[1] for res in results], "output_annual_principal_repayment.xlsx", "Principal"),
                file_name="output_annual_principal_repayment.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label="Download Annual Interest Repayment",
                data=create_excel_file([res[2] for res in results], "output_annual_interest_repayment.xlsx", "Interest"),
                file_name="output_annual_interest_repayment.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label="Download Outstanding Amount",
                data=create_excel_file([res[3] for res in results], "output_outstanding_amount.xlsx", "Outstanding"),
                file_name="output_outstanding_amount.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label="Download Current Liability",
                data=create_excel_file([res[4] for res in results], "output_current_liability.xlsx", "Liabilities"),
                file_name="output_current_liability.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label="Download Amortization Schedule",
                data=create_schedule_excel_file([res[0] for res in results], "output_amortization_schedule.xlsx"),
                file_name="output_amortization_schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Error calculating amortization: {str(e)}")

else:
    uploaded_file = st.file_uploader("Upload Excel file (up to 35 loans or more)", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
            
            required_columns = ['sr_no', 'bank_name_or_loan_no', 'loan_amount', 'interest_rate', 'loan_term', 'start_date']
            if not all(col in df.columns for col in required_columns):
                st.error("File must contain columns: Sr no, bank Name or loan no, Loan Amount, interest Rate, Loan term, Start Date")
            else:
                if len(df) > 35:
                    st.warning(f"File contains {len(df)} loans. Processing all, but performance may vary for very large datasets.")
                
                for idx, row in df.iterrows():
                    try:
                        sr_no = str(row['sr_no'])
                        loan_name = str(row['bank_name_or_loan_no'])
                        loan_amount_lakhs = float(row['loan_amount'])
                        interest_rate = float(row['interest_rate'])
                        loan_term_months = int(row['loan_term'])
                        start_date = row['start_date']
                        
                        if isinstance(start_date, str):
                            try:
                                start_date = datetime.strptime(start_date, '%m/%d/%Y').strftime('%Y-%m-%d')
                            except ValueError:
                                try:
                                    start_date = datetime.strptime(start_date, '%d/%m/%Y').strftime('%Y-%m-%d')
                                except ValueError:
                                    st.error(f"Invalid date format for Loan {sr_no}. Use MM/DD/YYYY or DD/MM/YYYY.")
                                    continue
                        elif isinstance(start_date, datetime):
                            start_date = start_date.strftime('%Y-%m-%d')
                        
                        payment_amount_lakhs = float(row['payment_amount']) if 'payment_amount' in row and pd.notna(row['payment_amount']) else None
                        
                        schedule_df, principal_pivot, interest_pivot, outstanding_pivot, liabilities_pivot = calculate_annual_metrics(
                            calculate_amortization(loan_amount_lakhs, interest_rate, loan_term_months, start_date, payment_amount_lakhs),
                            sr_no, loan_name, loan_amount_lakhs
                        )
                        results.append((schedule_df, principal_pivot, interest_pivot, outstanding_pivot, liabilities_pivot))
                    
                    except ValueError as e:
                        st.error(f"Error processing Loan {sr_no}: {str(e)}")
                        continue
                
                if results:
                    st.download_button(
                        label="Download Annual Principal Repayment",
                        data=create_excel_file([res[1] for res in results], "output_annual_principal_repayment.xlsx", "Principal"),
                        file_name="output_annual_principal_repayment.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="Download Annual Interest Repayment",
                        data=create_excel_file([res[2] for res in results], "output_annual_interest_repayment.xlsx", "Interest"),
                        file_name="output_annual_interest_repayment.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="Download Outstanding Amount",
                        data=create_excel_file([res[3] for res in results], "output_outstanding_amount.xlsx", "Outstanding"),
                        file_name="output_outstanding_amount.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="Download Current Liability",
                        data=create_excel_file([res[4] for res in results], "output_current_liability.xlsx", "Liabilities"),
                        file_name="output_current_liability.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="Download Amortization Schedule",
                        data=create_schedule_excel_file([res[0] for res in results], "output_amortization_schedule.xlsx"),
                        file_name="output_amortization_schedule.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    else:
        st.info("Please upload an Excel file to calculate the amortization schedule for multiple loans.")