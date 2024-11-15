import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from calendar import monthrange
from dateutil.relativedelta import relativedelta


def loan_amortization_sheet(row):
    compounding_period = {
        "Annually": 1,
        "Semi-Annually": 2,
        "Quarterly": 4,
        "Monthly": 12
    }

    payment_periods = {
        "Annually": {"periods_per_year": 1, "months_offset": 12, "day_offset": 0},
        "Semi-Annually": {"periods_per_year": 2, "months_offset": 6, "day_offset": 0},
        "Quarterly": {"periods_per_year": 4, "months_offset": 3, "day_offset": 0},
        "Bi-Monthly": {"periods_per_year": 6, "months_offset": 2, "day_offset": 0},
        "Monthly": {"periods_per_year": 12, "months_offset": 1, "day_offset": 0},
        "Semi-Monthly": {"periods_per_year": 24, "months_offset": 0, "day_offset": 15},
        "Bi-Weekly": {"periods_per_year": 26, "months_offset": 0, "day_offset": 14},
        "Weekly": {"periods_per_year": 52, "months_offset": 0, "day_offset": 7}
    }
    # Input values
    start_date = row["start_date"]
    original_principal = row["original_principal"]  # Amount
    # Number of months for the amortization
    amortization_term_months = row["amortization_term_months"]
    # Mortgage term (should be same as renewal_period)
    mortgage_term_months = row["mortgage_term_months"]
    interest_rate = row["interest_rate"]  # annual interest rate
    compounding_frequency = row["compounding_frequency"]  # Compounded monthly
    payment_frequency = row["payment_frequency"]  # payment intervals
    cpr = row["cpr"]  # constant prepayment rate

    def generate_payment_dates(start_date_str, frequency):
        # Parse the start date
        # datetime.strptime(start_date_str, "%Y-%m-%d")
        start_date = start_date_str

        # Get the payment frequency details
        freq_details = payment_periods[frequency]

        # Generate the payment dates
        payment_dates = []
        current_date = start_date
        for _ in range(freq_details["periods_per_year"]*((row["amortization_term_months"]//12)+1)):
            payment_dates.append(current_date.strftime("%Y-%m-%d"))
            if freq_details["months_offset"] > 0:
                # timedelta(days=30 * freq_details["months_offset"])
                current_date += relativedelta(
                    months=freq_details["months_offset"])
            else:
                # timedelta(days=freq_details["day_offset"])
                current_date += relativedelta(days=freq_details["day_offset"])

        return payment_dates

    def amortization(principal, amortization_term, interest, compounding_frequency, payment_frequency, cpr, mortgage_term=0):
        calculated_dict = dict()
        calculated_dict["compounding period"] = compounding_period[compounding_frequency]
        calculated_dict["periods per year"] = payment_periods[payment_frequency]["periods_per_year"]
        calculated_dict["interest rate per payment"] = ((1+(interest/calculated_dict["compounding period"]))**(
            calculated_dict["compounding period"]/calculated_dict["periods per year"]))-1
        if mortgage_term:
            calculated_dict["renewal period"] = (
                mortgage_term/12)*calculated_dict["periods per year"]
        else:
            calculated_dict["renewal period"] = 0
        calculated_dict["amortization period"] = (
            amortization_term/12)*calculated_dict["periods per year"]
        calculated_dict["payment per period"] = ((calculated_dict["interest rate per payment"])*principal)/(
            1-((1+(calculated_dict["interest rate per payment"]))**(-1*calculated_dict["amortization period"])))
        calculated_dict["smm"] = cpr/calculated_dict["periods per year"]
        calculated_dict["month offset"] = payment_periods[payment_frequency]["months_offset"]
        calculated_dict["day offset"] = payment_periods[payment_frequency]["day_offset"]
        return calculated_dict

    loanTape = row.name

    a = amortization(principal=original_principal, amortization_term=amortization_term_months, mortgage_term=mortgage_term_months,
                     interest=interest_rate, compounding_frequency=compounding_frequency, payment_frequency=payment_frequency, cpr=cpr)

    # Compounding frequency per year
    compounding_period = a["compounding period"]
    periods_per_year = a["periods per year"]
    # interest per period
    interest_rate_perpayment = a["interest rate per payment"]
    renewal_period = a["renewal period"]  # Annual renewal
    amortization_period = a["amortization period"]
    payment_per_period = a["payment per period"]  # Fixed payment amount
    smm = a["smm"]  # (Single Monthly Mortality rate)
    month_offset = a["month offset"]  # Payment month offset
    day_offset = a["day offset"]  # Day offset

    # Convert the start date to a datetime object
    # start_date = datetime.strptime(start_date, "%d-%m-%Y")

    # Initialize an empty list to store the rows for the DataFrame
    amortization_schedule = []

    # Set the initial values for period 0
    period = 0
    date = start_date
    opening_balance = original_principal
    new_origination = original_principal
    closing_balance = new_origination

    dates = generate_payment_dates(
        start_date_str=start_date, frequency=payment_frequency)
    # Loop over the periods
    for period in range(int(a["amortization period"]) + 1):  # +1 to include period 0
        # Calculate the interest for the period
        interest = (opening_balance * interest_rate_perpayment)

        # Payment is the minimum of payment_per_period and the remaining balance (including interest)
        if opening_balance > payment_per_period:
            payment = (payment_per_period)
        else:
            payment = (opening_balance + interest)

        # Calculate principal paid for the period
        principal = payment - interest

        # Calculate prepayment
        if opening_balance - principal > 0:
            prepayment = (opening_balance * smm)
        else:
            prepayment = 0

        if not (period):
            opening_balance = 0
            payment = 0
            interest = 0
            principal = 0
            prepayment = 0

        else:
            new_origination = 0

        maturity = 0
        # Apply prepayment, maturity, and calculate closing balance
        if period < renewal_period or mortgage_term_months == 0:
            # For the first 'mortgage_term_months', no maturity (i.e., closing_balance is just reduced by prepayment and principal)
            closing_balance = opening_balance - prepayment - principal + new_origination
        elif (period >= renewal_period and renewal_period > 0) or period >= amortization_period:
            # After mortgage term ends, apply maturity
            # Remaining balance after regular payments
            maturity = opening_balance - principal - prepayment
            closing_balance = 0  # After maturity, the loan is fully paid off

        # Prepare the data for this period
        row = {
            "period": period,
            "date": dates[period],  # date,#.strftime("%d-%m-%Y"),
            "opening_balance": opening_balance,
            "payment": payment,
            "interest": interest,
            "principal": principal,
            "prepayment": prepayment,
            "new_origination": new_origination,
            "maturity": maturity if (period >= renewal_period and mortgage_term_months != 0) else 0,
            "closing_balance": closing_balance
        }

        # Add the row to the amortization schedule
        amortization_schedule.append(row)

        # Update the opening balance for the next period
        # Opening balance of the next period is the closing balance of the current period
        opening_balance = closing_balance

        # Increment the date
        if (period >= renewal_period and renewal_period > 0) or period >= amortization_period:
            break
        # date = process_formula(period+1, month_offset, day_offset, date)

    # Create the DataFrame from the amortization schedule list
    df_amortization = pd.DataFrame(amortization_schedule)
    df_amortization["date"] = pd.to_datetime(
        df_amortization["date"]).dt.strftime('%d-%m-%Y')

    #### DAILY SCHEDULE#####
    df_amortization['date'] = pd.to_datetime(
        df_amortization['date'], format="%d-%m-%Y")

    # Calculate daily interest rate
    # Monthly rate, assuming monthly compounding
    interest_rate_per_payment = 0.05 / 12
    days_in_month = 30  # Adjust for actual days in month if needed, or calculate dynamically
    daily_interest_rate = interest_rate_per_payment / days_in_month

    # To store the daily amortization data
    daily_amortization_records = []
    # Iterate over each row of the periodic amortization table
    for index, row in df_amortization.iterrows():
        # Get the payment date and details from the periodic table
        period = row["period"]
        payment_date = row['date']
        opening_balance = row['opening_balance']
        payment = row['payment']
        interest = row['interest']
        principal = row['principal']
        prepayment = row["prepayment"]
        new_origination = row["new_origination"]
        maturity = row["maturity"]
        closing_balance = row['closing_balance']

        # Calculate the interest for the days before the payment date
        current_date = start_date
        while current_date < payment_date:
            # Calculate daily interest based on the opening balance
            daily_interest = opening_balance * daily_interest_rate
            # Add daily record
            daily_amortization_records.append({
                "period": period-1,
                'date': current_date,
                'opening_balance': opening_balance,
                'payment': 0,
                'interest': 0,  # daily_interest,
                'principal': 0,
                'prepayment': 0,
                "new_origination": 0,
                'maturity': 0,
                'closing_balance': opening_balance
            })

            # Move to the next day
            current_date += timedelta(days=1)

        # Now, we have reached the payment day, record the payment details
        # Record the payment for the current day
        daily_amortization_records.append({
            'period': period,
            'date': payment_date,
            'opening_balance': opening_balance,
            'payment': payment,
            'interest': interest,
            'principal': principal,
            'prepayment': prepayment,
            "new_origination": new_origination,
            "maturity": maturity,
            'closing_balance': closing_balance
        })

        # Update start_date for the next period
        start_date = payment_date + timedelta(days=1)
    # Create a new DataFrame from the records
    daily_amortization_df = pd.DataFrame(daily_amortization_records)
    daily_amortization_df["date"] = pd.to_datetime(
        daily_amortization_df["date"]).dt.strftime('%d-%m-%Y')

    ##### MONTHLY####
    # Convert 'date' column to datetime format
    daily_amortization_df['date'] = pd.to_datetime(
        daily_amortization_df['date'], format='%d-%m-%Y')

    # Group by year and month, aggregate the values for each month
    df_monthly = daily_amortization_df.groupby(daily_amortization_df['date'].dt.to_period('M')).agg({
        'opening_balance': 'first',  # The opening balance is the first value in the month
        'payment': 'sum',            # Sum of payments in the month
        'interest': 'sum',           # Sum of interest in the month
        'principal': 'sum',          # Sum of principal in the month
        'prepayment': 'sum',         # Sum of prepayments in the month
        'new_origination': 'sum',    # Sum of new origination in the month
        'maturity': 'sum',           # Sum of maturity in the month
        'closing_balance': 'last'    # The closing balance is the last value in the month
    }).reset_index()

    # Convert period back to datetime format (e.g., '2024-02' to '01-02-2024')
    df_monthly['date'] = df_monthly['date'].dt.to_timestamp()
    df_monthly["date"] = pd.to_datetime(
        df_monthly["date"]).dt.strftime('%d-%m-%Y')
    df_amortization["date"] = pd.to_datetime(
        df_amortization["date"]).dt.strftime('%d-%m-%Y')
    daily_amortization_df["date"] = pd.to_datetime(
        daily_amortization_df["date"]).dt.strftime('%d-%m-%Y')

    filename = "individual amortization tables/LoanTape " + \
        str(loanTape)+".xlsx"
    # Create a Pandas Excel writer using a context manager
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write each DataFrame to a different sheet
        df_amortization.round(2).to_excel(
            writer, sheet_name='Periodic Amortization Table', index=False)
        daily_amortization_df.round(2).to_excel(
            writer, sheet_name='Daily Amortization Table')
        df_monthly.round(2).to_excel(
            writer, sheet_name='Monthly Amortization Table', index=False)

    print("Ammortion tables for LoanTape " + str(loanTape) +
          " have been written to '" + filename+"'")


if "__main__":
    loan_data = pd.read_excel("sample_dataset.xlsx")

    loan_data.apply(loan_amortization_sheet, axis=1)

    daily_amortization = []
    monthly_amortization = []

    for i in range(len(loan_data)):
        daily_amortization.append(pd.read_excel(
            "individual amortization tables/LoanTape " + str(i)+".xlsx", sheet_name='Daily Amortization Table'))
        monthly_amortization.append(pd.read_excel(
            "individual amortization tables/LoanTape " + str(i)+".xlsx", sheet_name='Monthly Amortization Table'))

    consolidated_daily_df = pd.DataFrame()
    consolidated_monthly_df = pd.DataFrame()

    consolidated_daily_df = pd.concat(daily_amortization)
    # consolidated_daily_df.to_excel("daily test.xlsx")
    consolidated_daily_df = consolidated_daily_df.groupby(
        'date').sum().reset_index()
    # Convert the 'date' column to datetime format
    consolidated_daily_df['date'] = pd.to_datetime(
        consolidated_daily_df['date'], format='%d-%m-%Y')

    # Sort the dataframe by the 'date' column
    consolidated_daily_df = consolidated_daily_df.sort_values(by='date')

    consolidated_monthly_df = pd.concat(monthly_amortization)
    # consolidated_daily_df.to_excel("daily test.xlsx")
    consolidated_monthly_df = consolidated_monthly_df.groupby(
        'date').sum().reset_index()
    # Convert the 'date' column to datetime format
    consolidated_monthly_df['date'] = pd.to_datetime(
        consolidated_monthly_df['date'], format='%d-%m-%Y')

    # Sort the dataframe by the 'date' column
    consolidated_monthly_df = consolidated_monthly_df.sort_values(by='date')
    consolidated_daily_df["date"] = consolidated_daily_df["date"].dt.strftime(
        '%d-%m-%Y')
    consolidated_monthly_df["date"] = consolidated_monthly_df["date"].dt.strftime(
        '%d-%m-%Y')
    consolidated_daily_df = consolidated_daily_df.reset_index()
    # consolidated_daily_df.rename(columns={'index': 'day'}, inplace=True)
    consolidated_daily_df.drop(
        columns=["Unnamed: 0", "period", "index"], inplace=True)
    # consolidated_monthly_df.drop(columns=["Unnamed: 0"], inplace=True)
    consolidated_daily_df.rename(columns={'index': 'day'}, inplace=True)

    filename = "Consolidated Tables.xlsx"
    # # Create a Pandas Excel writer using a context manager
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write each DataFrame to a different sheet
        consolidated_daily_df.round(2).to_excel(
            writer, sheet_name='Consolidated Daily Table')
        consolidated_monthly_df.round(2).to_excel(
            writer, sheet_name='Consolidated Monthly Table', index=False)

    print("Consolidated Amortization tables stored in 'Consolidated Tables.xlsx'")
