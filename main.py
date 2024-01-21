import yfinance as yf
import pandas as pd


def get_stock_prices(tickers, start_date, end_date):
    # Download historical stock data for multiple tickers
    data = yf.download(tickers, start=start_date, end=end_date)

    # Extract closing prices
    prices = data['Close']

    # Reset index to make it a flat DataFrame
    prices.reset_index(inplace=True)

    # Convert the 'Date' column to TimedeltaIndex
    prices['Date'] = pd.to_datetime(prices['Date'])

    # Convert the 'Date' column to TimedeltaIndex
    prices['Date'] = pd.to_timedelta(prices['Date'] - prices['Date'].min())

    return prices


def save_to_excel(prices, output_file):
    # Create a copy of the file to avoid a warning
    prices_copy = prices.copy()

    # Save DataFrame to Excel using .loc
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        sheet_name = 'Stock Prices'
        prices_copy.to_excel(writer, sheet_name=sheet_name, index=False)

        # Use .loc to set the sheet name in the Excel file
        writer.sheets[sheet_name].title = sheet_name

    print(f"Data saved to {output_file}")


# Example usage
ticker_symbols = ['AAPL', 'GOOGL', 'MSFT']  # Apple, Google and Microsoft
start_date = '2023-01-21'
end_date = '2024-01-21'
output_file = '/Users/christophermumma/Documents/Python Projects/StockPriceTracker/Stock Prices/stock_prices.xlsx'

stock_prices = get_stock_prices(ticker_symbols, start_date, end_date)
save_to_excel(stock_prices, output_file)
