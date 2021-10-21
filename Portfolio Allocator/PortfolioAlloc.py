import math
import numpy as np
import pandas as pd
import requests
import xlsxwriter
from secrets import IEX_CLOUD_API_TOKEN


def chunks(lst, n):
    # Produces n-sized chunks from a list
    for i in range(0, len(lst), n):
        yield lst[i:i+n]


portfolio_size = 10000000.0 
stocks = pd.read_csv('sp_500_stocks.csv')
fund_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
fund_df = pd.DataFrame(columns=fund_columns)
counter = 0

for stock in stocks['Ticker']:
    api_url = f'https://sandbox.iexapis.com/stable/stock/{stock}/quote?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(api_url).json()
    fund_df = fund_df.append(
                            pd.Series([stock, data['latestPrice'], data['marketCap'], 'N/A'],
                                        index=fund_columns
                                    ),
                            ignore_index=True
                            )
    counter += 1
    print(f'{counter} of {len(stocks)} stocks downloaded', end='\r')

symbol_batches = list(chunks(stocks['Ticker'], 100))
symbols_strings = []

for i in symbol_batches:
    symbols_strings.append(','.join(i))

for i in symbols_strings:
    batch_api_url = f"https://sandbox.iexapis.com/stable/stock/market/batch?symbols={i}&types=quote&token={IEX_CLOUD_API_TOKEN}"
    data = requests.get(batch_api_url).json() 
    
    for stock in i.split(','):
        fund_df = fund_df.append(
            pd.Series(
                [
                    stock,
                    data[stock]['quote']['latestPrice'],
                    data[stock]['quote']['marketCap'],
                    'N/A'
                ],
                index=fund_columns
            ),
            ignore_index=True
        )

index_market_val = fund_df['Market Capitalization'].sum()

fund_df['Number of Shares to Buy'] = (portfolio_size * fund_df['Market Capitalization'] / index_market_val) // fund_df['Stock Price'] # no fractional shares

writer = pd.ExcelWriter('SP500RecIndex.xlsx', engine='xlsxwriter')
fund_df.to_excel(writer, 'Recommended Trades', index=False)

# Formatting style of spreadsheet
bg_color = "#0A0A23"
font_color = "#FFFFFF"
font_name = 'Consolas'
string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'font_name': font_name,
        'bg_color': bg_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_name': font_name,
        'font_color': font_color,
        'bg_color': bg_color,
        'border': 1
    }
)

int_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_name': font_name,
        'font_color': font_color,
        'bg_color': bg_color,
        'border': 1
    }
)

column_formats = {
    'A': [fund_columns[0], string_format],
    'B': [fund_columns[1], dollar_format],
    'C': [fund_columns[2], dollar_format],
    'D': [fund_columns[3], int_format]
}

for column in column_formats:
    writer.sheets['Recommended Trades'].set_column(f"{column}:{column}", 18, column_formats[column][1])
    writer.save()