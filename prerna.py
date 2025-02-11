import requests
import pandas as pd
import openpyxl
from openpyxl import Workbook
import time
import datetime


def fetch_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': 50,  
        'page': 1
    }
    
    response = requests.get(url, params=params)
    return response.json()


def analyze_data(data):
    df = pd.DataFrame(data)
    

    df = df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]


    top_5 = df.sort_values(by='market_cap', ascending=False).head(5)
    

    average_price = df['current_price'].mean()
    

    max_change = df.loc[df['price_change_percentage_24h'].idxmax()]
    min_change = df.loc[df['price_change_percentage_24h'].idxmin()]
    
    return df, top_5, average_price, max_change, min_change


def update_excel(df, top_5, average_price, max_change, min_change):

    try:
        wb = openpyxl.load_workbook('crypto_data.xlsx')
        sheet = wb.active
    except FileNotFoundError:
        wb = Workbook()
        sheet = wb.active
        sheet.append(["Name", "Symbol", "Current Price (USD)", "Market Cap", "24h Volume", "Price Change (24h)", "Timestamp"])


    for _, row in df.iterrows():
        sheet.append([row['name'], row['symbol'], row['current_price'], row['market_cap'], 
                      row['total_volume'], row['price_change_percentage_24h'], datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    

    sheet.append(['Top 5 Cryptocurrencies'])
    for _, row in top_5.iterrows():
        sheet.append([row['name'], row['symbol'], row['market_cap']])
    
    sheet.append(['Average Price of Top 50', average_price])
    sheet.append(['Max 24h Price Change', max_change['name'], max_change['symbol'], max_change['price_change_percentage_24h']])
    sheet.append(['Min 24h Price Change', min_change['name'], min_change['symbol'], min_change['price_change_percentage_24h']])
    

    wb.save('crypto_data.xlsx')


def main():
    while True:
        data = fetch_data()
        df, top_5, average_price, max_change, min_change = analyze_data(data)
        update_excel(df, top_5, average_price, max_change, min_change)
        print(f"Data updated at {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        time.sleep(300)  


if __name__ == "__main__":
    main()
