# Stock Price Lookup

This project is a Python script that allows you to look up historical stock prices for a specified ticker and date range. It retrieves daily stock data (including opening and closing prices) and computes the daily gain/loss using the formula `(Closing Price / Opening Price - 1)`. The results are then saved to an Excel file for further analysis.

## Features

- **Historical Data Retrieval:** Fetch daily stock prices using the [yfinance](https://pypi.org/project/yfinance/) library.
- **Data Calculation:** Compute daily gain/loss based on opening and closing prices.
- **Excel Export:** Save the results to an Excel file with the following columns:
  - Stock Symbol
  - Stock Name
  - Date
  - Opening Price
  - Closing Price
  - gain/loss
- **Timezone Handling:** Automatically removes timezone information from datetime values to ensure compatibility with Excel.
- **Interactive Confirmation:** Displays the fetched stock symbol and company name for user confirmation before processing.

## Prerequisites

- Python 3.x

### Required Python Libraries

- [yfinance](https://pypi.org/project/yfinance/)
- [pandas](https://pypi.org/project/pandas/)
- [openpyxl](https://pypi.org/project/openpyxl/)

## Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/your-username/stock-price-lookup.git
   cd stock-price-lookup
   ```

2. **Create a Virtual Environment (Optional but Recommended):**

   ```bash
   python -m venv env
   ```

   Activate the virtual environment:
   - **Windows:** `env\Scripts\activate`
   - **macOS/Linux:** `source env/bin/activate`

3. **Install the Required Packages:**

   ```bash
   pip install yfinance pandas openpyxl
   ```

## Usage

Run the Python script by executing:

```bash
python find_stock_prices.py
```

When prompted, enter:
- The stock ticker (e.g., `AAPL`).
- The start date (in `YYYY-MM-DD` format).
- The end date (in `YYYY-MM-DD` format).

The script will display the stock symbol and company name for confirmation. Once confirmed, it will fetch the data, compute the gain/loss, and export the results to an Excel file (e.g., `AAPL_stock_data.xlsx`).

### Example Interaction

```
Enter the stock ticker: AAPL
Enter the start date (YYYY-MM-DD): 2023-01-01
Enter the end date (YYYY-MM-DD): 2023-01-31

Stock Symbol: AAPL
Stock Name: Apple Inc.
Is this correct? (Y/N): Y

Data has been saved to 'AAPL_stock_data.xlsx'.
```

## Contributing

Contributions are welcome! Feel free to open an issue or submit a pull request with improvements, bug fixes, or new features.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Disclaimer

This script uses data provided by Yahoo Finance via the [yfinance](https://pypi.org/project/yfinance/) library. The data is provided "as is" without any warranty. Always verify financial data with official sources before making any investment decisions.

---

Feel free to modify this `README.md` as needed to fit your project's specifics.
