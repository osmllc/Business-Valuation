import pandas as pd
from alpha_vantage.fundamentaldata import FundamentalData
import os

# Your Alpha Vantage API key
api_key = 'XXXX'

# Initialize the Alpha Vantage FundamentalData object
fd = FundamentalData(key=api_key, output_format='pandas')

# Define the list of ticker symbols (reduced to avoid hitting rate limits)
tickers = ['JPM', 'NVDA','UBER','XOM','HII','OR','EL','LMT']  # Reduce the number of tickers to avoid rate limit

# Define the filename
filename = 'stocks_financial_data.xlsx'

# Ensure the file is not open and has write permission
if os.path.exists(filename):
    try:
        os.rename(filename, filename)  # Try to rename the file to check if it's open
    except OSError as e:
        print(f"Error: {e.strerror}. Please close the file if it's open and try again.")
        exit()

# Create an Excel writer object
with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    any_data_written = False  # Flag to check if any data is written to the file

    for ticker in tickers:
        try:
            # Fetch the income statement data
            income_statement, _ = fd.get_income_statement_annual(ticker)
            income_statement = income_statement.set_index('fiscalDateEnding')
            income_statement.index = pd.to_datetime(income_statement.index)
            income_statement = income_statement[income_statement.index >= '2012-01-01']

            # Fetch the balance sheet data
            balance_sheet, _ = fd.get_balance_sheet_annual(ticker)
            balance_sheet = balance_sheet.set_index('fiscalDateEnding')
            balance_sheet.index = pd.to_datetime(balance_sheet.index)
            balance_sheet = balance_sheet[balance_sheet.index >= '2012-01-01']

            # Fetch the cash flow data (for dividends per share and shares outstanding)
            cash_flow, _ = fd.get_cash_flow_annual(ticker)
            cash_flow = cash_flow.set_index('fiscalDateEnding')
            cash_flow.index = pd.to_datetime(cash_flow.index)
            cash_flow = cash_flow[cash_flow.index >= '2012-01-01']

            # Debugging: print available columns
            print(f"Available columns in income statement for {ticker}: {income_statement.columns}")
            print(f"Available columns in balance sheet for {ticker}: {balance_sheet.columns}")
            print(f"Available columns in cash flow for {ticker}: {cash_flow.columns}")

            # Select the required financial metrics
            financial_data = pd.DataFrame({
                'Revenue': income_statement.get('totalRevenue', pd.Series(dtype='float64')),
                'Net Income': income_statement.get('netIncome', pd.Series(dtype='float64')),
                'EPS': income_statement.get('reportedEPS', pd.Series(dtype='float64')),
                'Shareholder Equity': balance_sheet.get('totalShareholderEquity', pd.Series(dtype='float64')),
                'Dividends per Share': cash_flow.get('dividendPayout', pd.Series(dtype='float64')),
                'Shares Outstanding': balance_sheet.get('commonStockSharesOutstanding', pd.Series(dtype='float64'))
            })

            # Convert financial data to numeric and handle missing values
            financial_data = financial_data.apply(pd.to_numeric, errors='coerce').fillna(0)

            # Scale financial figures to millions of dollars, except for EPS, Dividends per Share, and Shares Outstanding
            for column in ['Revenue', 'Net Income', 'Shareholder Equity']:
                financial_data[column] = financial_data[column] / 1e6

            # Transpose the DataFrame to make the table horizontal
            financial_data = financial_data.T

            # Reorder the columns to start from 2012 and end with the latest year available
            financial_data = financial_data[sorted(financial_data.columns)]

            # Write the DataFrame to a sheet in the Excel file
            financial_data.to_excel(writer, sheet_name=ticker)
            any_data_written = True

            # Display the results for the current ticker
            print(f"\nFinancial Data for {ticker}:\n", financial_data)

        except ValueError as e:
            print(f"Error for {ticker}: {e}")
            continue

    if not any_data_written:
        # Ensure at least one sheet is visible
        sheet = writer.book.create_sheet("Sheet1")
        writer.book.active = writer.book.index(sheet)

# Display the combined results
print("\nAll data has been written to 'stocks_financial_data.xlsx'")
