import yfinance as yf
import pandas as pd

def main():
    # Get user input
    ticker_input = input("Enter the stock ticker: ").upper().strip()
    start_date = input("Enter the start date (YYYY-MM-DD): ").strip()
    end_date = input("Enter the end date (YYYY-MM-DD): ").strip()

    # Get ticker object from yfinance and retrieve stock info
    ticker = yf.Ticker(ticker_input)
    info = ticker.info

    # Retrieve the company name (using shortName or longName as available)
    stock_name = info.get('shortName') or info.get('longName') or "N/A"

    # Show the symbol and name for confirmation
    print(f"\nStock Symbol: {ticker_input}")
    print(f"Stock Name: {stock_name}")
    confirm = input("Is this correct? (Y/N): ").strip().lower()
    if confirm != 'y':
        print("Exiting the program.")
        return

    # Download historical data for the specified date range
    df = ticker.history(start=start_date, end=end_date)

    # Check if data is available
    if df.empty:
        print("No data found for the specified date range.")
        return

    # Reset the index to bring the Date into a column
    df = df.reset_index()

    # Remove timezone information from the Date column if present
    if pd.api.types.is_datetime64tz_dtype(df['Date']):
        df['Date'] = df['Date'].dt.tz_localize(None)

    # Compute gain/loss using the provided formula: (Closing Price/Opening Price - 1)
    df['gain/loss'] = (df['Close'] / df['Open'] - 1)

    # Add stock symbol and stock name columns
    df['Stock Symbol'] = ticker_input
    df['Stock Name'] = stock_name

    # Rename columns to match the required output format
    df.rename(columns={'Open': 'Opening Price', 'Close': 'Closing Price'}, inplace=True)

    # Rearrange columns as desired
    final_df = df[['Stock Symbol', 'Stock Name', 'Date', 'Opening Price', 'Closing Price', 'gain/loss']]

    # Export the DataFrame to an Excel file
    file_name = f"{ticker_input}_stock_data.xlsx"
    final_df.to_excel(file_name, index=False)
    print(f"\nData has been saved to '{file_name}'.")

if __name__ == "__main__":
    main()
