import yfinance as yf
import pandas as pd

# Define the list of ticker symbols
tickers = ['JPM', 'NVDA', 'TSLA', 'CRM', 'JNJ', 'KO', 'LPX', 'MCO', 'T', 'V', 'VZ', 'WMT']  # Add more tickers as needed

# Define the start and end dates
start_date = '2012-01-01'
end_date = '2024-07-11'  # You can set this to today's date or any end date you prefer

# Create an Excel writer object
with pd.ExcelWriter('stocks_yearly_data_horizontal.xlsx', engine='openpyxl') as writer:
    for ticker in tickers:
        # Initialize the Ticker object
        stock = yf.Ticker(ticker)

        # Fetch the historical data
        historical_data = stock.history(start=start_date, end=end_date)

        # Extract the year from the date index
        historical_data['Year'] = historical_data.index.year

        # Initialize dictionaries to hold the financial metrics for each year
        max_prices = {}
        min_prices = {}
        dividends = {}

        # Loop through each year to find the required data
        for year in historical_data['Year'].unique():
            yearly_data = historical_data[historical_data['Year'] == year]
            max_prices[year] = yearly_data['Close'].max()
            min_prices[year] = yearly_data['Close'].min()

            # Assuming dividends are provided quarterly, sum them up for the year
            yearly_dividends = stock.dividends[stock.dividends.index.year == year]
            dividends[year] = yearly_dividends.sum() if not yearly_dividends.empty else float('nan')

        # Create a DataFrame with the financial metrics
        yearly_data = pd.DataFrame({
            'Min Price': min_prices,
            'Max Price': max_prices,
            'Dividends': dividends,
        })

        # Transpose the DataFrame to make the table horizontal
        yearly_data = yearly_data.T

        # Reorder the columns to start from 2012
        yearly_data = yearly_data[sorted(yearly_data.columns)]

        # Reorder the rows to have Min Price on top and Max Price on bottom
        yearly_data = yearly_data.loc[['Min Price', 'Max Price', 'Dividends']]

        # Write the DataFrame to a sheet in the Excel file
        yearly_data.to_excel(writer, sheet_name=ticker)

        # Display the results for the current ticker
        print(f"\nData for {ticker}:\n", yearly_data)

# Display the combined results
print("\nAll data has been written to 'stocks_yearly_data_horizontal.xlsx'")
