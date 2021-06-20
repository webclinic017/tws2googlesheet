import pygsheets
import requests_cache
import pandas_datareader as pdr
import datetime as dt
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
client = pygsheets.authorize(service_file='pythonsheet.json')

sheet = client.open('Haupttabelle')
wks2 = sheet.worksheet_by_title('Untertabelle')
rows = int(wks2.cell((1,1)).value)

get_date = lambda : datetime.utcnow().strftime('%d-%m-%Y')

expire_after = dt.timedelta(days=1)
session = requests_cache.CachedSession(cache_name='cache', backend='sqlite', expire_after=expire_after) # Cache, damit nicht alles jedesmal neu von Yahoo geladen wird

df = pd.DataFrame()

def ticker():
    global df
    ticker = wks2.cell((i, 2)).value
    rsi = pdr.get_data_yahoo(ticker, dt.datetime(2021, 5, 1), dt.datetime.now(), session=session)
    sma = pdr.get_data_yahoo(ticker, dt.datetime(2020, 5, 1), session=session)
    stoch = pdr.get_data_yahoo(ticker, dt.datetime(2021, 5, 1), dt.datetime.now(), session=session)
    if isinstance(rsi, pd.core.frame.DataFrame): 
        delta = rsi['Close'].diff()
        up = delta.clip(lower=0)
        down = -1*delta.clip(upper=0)
        ema_up = up.ewm(com=13, adjust=False).mean()
        ema_down = down.ewm(com=13, adjust=False).mean()
        rs = ema_up/ema_down
        rsi['RSI'] = 100-(100/(1+rs))
        rsi_value = round(rsi.iloc[-1]['RSI'],2)
    else:
        rsi_value = '0'
    if isinstance(stoch, pd.core.frame.DataFrame):
        stoch['14-high'] = stoch['High'].rolling(14).max()
        stoch['14-low'] = stoch['Low'].rolling(14).min()
        stoch['%K'] = (stoch['Close'] - stoch['14-low'])*100/(stoch['14-high'] - stoch['14-low'])
        stoch['%D'] = stoch['%K'].rolling(3).mean()
        stoch2 = round(stoch.iloc[-1]['%D'],2)
    else:
        stoch2 = '0'
    if isinstance(sma, pd.core.frame.DataFrame):
        sma['SMA10'] = sma['Close'].rolling(10).mean()
        sma['SMA50'] = sma['Close'].rolling(50).mean()
        sma['SMA200'] = sma['Close'].rolling(200).mean()
        sma_value = round(sma.iloc[-1]['SMA200'],2)
    else:
        sma_value = '0'
    print(ticker, rsi_value, sma_value, stoch2)
    liste = [[sma_value, rsi_value, stoch2]]
    df2 = pd.DataFrame(liste)
    df = df.append(df2, ignore_index=True)

def main():
    try:
        ticker()
    except Exception as e:
        print(e)
        pass 

if __name__ == "__main__":
    for i in range(2, rows+2):
        main()

wks2.set_dataframe(df, 'O2', copy_head=False) # Upload zu Google in Spalte O,P,Q
