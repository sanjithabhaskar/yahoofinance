import requests
from bs4 import BeautifulSoup
import pandas as pd
from typing import Any, Dict

def get_data(ticker_symbol: Any) -> Dict[str, Any]:
    """
    Get stock data for a given ticker symbol from Yahoo Finance.
    Parameters:
    - ticker_symbol (Any): Ticker symbol of the stock.
    Returns:
    - Dict[str, Any]: Dictionary containing stock data.
    """
    print(f'Getting stock data for {ticker_symbol}...')

    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:122.0) Gecko/20100101 Firefox/122.0'
    }

    url = f'https://finance.yahoo.com/quote/{ticker_symbol}'
    r = requests.get(url, headers=headers)
    soup = BeautifulSoup(r.text, 'html.parser')

    try:
        stock_data = {'Stock Name': ticker_symbol}
        
        # Extract price, price change, and percentage change
        price_tag = soup.find('span', {'data-testid': 'qsp-price'})
        price_change_tag = soup.find('span', {'data-testid': 'qsp-price-change'})
        percent_change_tag = soup.find('span', {'data-testid': 'qsp-price-change-percent'})
        
        stock_data['Price'] = price_tag.text.strip() if price_tag else 'N/A'
        stock_data['Price Change'] = price_change_tag.text.strip() if price_change_tag else 'N/A'
        stock_data['Change Percentage'] = percent_change_tag.text.strip() if percent_change_tag else 'N/A'
        
        # Extract additional stock data
        statistics_div = soup.find('div', {'data-testid': 'quote-statistics'})
        if statistics_div:
            statistics_items = statistics_div.find_all('li', class_='yf-1jj98ts')
            for item in statistics_items:
                label = item.find('span', class_='label yf-1jj98ts')
                value = item.find('span', class_='value yf-1jj98ts')
                if label and value:
                    label_text = label.get('title', label.text).strip()
                    value_text = value.text.strip()
                    stock_data[label_text] = value_text
        
        return stock_data
    except Exception as e:
        return {'Stock Name': ticker_symbol, 'error': str(e)}

# List of ticker symbols
ticker_symbols = ['AAPL', 'GOOGL', 'AMZN', 'MSFT']  # Add your stock symbols here

# List to store data for each stock
all_stock_data = []

for ticker_symbol in ticker_symbols:
    stock_data = get_data(ticker_symbol)
    if 'error' in stock_data:
        print(f"Error retrieving data for {ticker_symbol}: {stock_data['error']}")
    else:
        all_stock_data.append(stock_data)

# Create DataFrame from all stock data
df = pd.DataFrame(all_stock_data)

# Save data to Excel
EXCEL_FILE_PATH = 'Bhaskar_stock_profile_data.xlsx'
df.to_excel(EXCEL_FILE_PATH, index=False)

print('\nDone! Data for 4 stocks saved to Excel file.')
