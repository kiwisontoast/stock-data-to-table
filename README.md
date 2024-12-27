# Stock-Data-to-Table
StockDatatoTable generates and compares financial metrics across multiple stocks. This tool displays stock performance and fundamental metrics in an organized table format using real-time data from Yahoo Finance.

## Key Features
- Multi-stock comparison with support for comma-separated ticker inputs
- Real-time data fetching using yfinance API
- Financial metrics including:
    - Trailing and Forward P/E
    - Market Cap
    - Beta
    - Price to Book
    - EV/EBITDA
    - EPS (TTM)
    - Revenue
    - Net Income
    - Profit Margin
    - Dividend Yield
- Responsive UI with:
    - Dark/Light theme toggle
    - Metric selection
    - Expandable table 
- Export Capabilities
    - Copy to clipboard functionality
    - Excel export feature
- Dynamic sizing 
- Error handling and user feedback 

## Dependencies
- This project requires:
    ~ tkinter
    ~ pandas
    ~ yfinance
    ~ sv_ttk
    ~ datetime

## Usage
- Enter stock tickers in the input field (comma-separated)
- Select desired metrics using the checkboxes
- Click "Generate Table" to fetch and display data 
- Use export options to:
    - Copy data to clipboard
    - Export to excel (saves as stock_data.xlsx)
- Toggle between light and dark mode as desired using the 
"Switch Theme" button.

## Warnings
**Data Retrieval**
- Ensure valid stock tickers are entered
- Some metrics may show "N/A" if they are unavailable
**Performance**
- Processing time increases with the number of stocks and selected metrics
- Large datasets may take longer to fetch and display
**Files**
- Ensure no other process is using "stock_data.xlsx" when exporting

## Authors 
Dev Shroff & Krishnika Anandan




