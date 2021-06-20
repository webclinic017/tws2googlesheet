import pygsheets
import pandas_datareader as pdr
import datetime as dt
import json
import requests
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
client = pygsheets.authorize(service_file='pythonsheet.json')

sheet = client.open('Hauptabelle')
wks2 = sheet.worksheet_by_title('Untertabelle')
rows = int(wks2.cell((1,1)).value)

get_date = lambda : datetime.utcnow().strftime('%d-%m-%Y')

df = pd.DataFrame()

def ticker():
    global df
    ticker = wks2.cell((i, 2)).value
    lhs_url = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/'
    rhs_url = '?formatted=true&crumb=swg7qs5y9UP&lang=en-US&region=US&' \
              'modules=upgradeDowngradeHistory,recommendationTrend,' \
              'financialData,earningsHistory,earningsTrend,industryTrend&' \
              'corsDomain=finance.yahoo.com'
    url =  lhs_url + ticker + rhs_url
    req = requests.get(url)
    rating = req.json()['quoteSummary']['result'][0]['financialData']['recommendationKey'] 
    url_vol = 'https://query2.finance.yahoo.com/v10/finance/quoteSummary/'
    url_vol2 = '?formatted=true&crumb=8ldhetOu7RJ&lang=en-US&region=US&modules=summaryDetail&corsDomain=finance.yahoo.com'
    vol_url = url_vol + ticker + url_vol2
    req_vol = requests.get(vol_url)
    volume = req_vol.json()['quoteSummary']['result'][0]['summaryDetail']['volume']['raw']
    print(ticker, rating, volume)
    liste = [[rating, volume]]
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

wks2.set_dataframe(df, 'M2', copy_head=False) # Upload zu Google in Spalte M und N
