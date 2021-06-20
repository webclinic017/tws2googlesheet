import xml.etree.ElementTree as ET
import pygsheets
from ib_insync import *
import gspread
from datetime import datetime
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
client = pygsheets.authorize(service_file='pythonsheet.json') # Datei mit den Zugangsdaten

sheet = client.open('Haupttabelle') # hier Name der Haupttabelle eintragen
wks2 = sheet.worksheet_by_title('Untertabelle') # hier Name der Untertabelle eintragen
rows = int(wks2.cell((1,1)).value) # Benutzte Reihen in Tabelle

ib = IB()
ib.connect('127.0.0.1', 7496, clientId=3) # Verbindung zur TWS, client-ID darf pro Script nur 1x vergeben sein

df = pd.DataFrame()

def ticker():
    global df
    ticker = wks2.cell((i, 2)).value # Ticker-Symbol aus Google-Tabelle
    contract = Stock(ticker, 'SMART', 'USD')
    ib.reqMarketDataType(1) # 1: Live, 2: Frozen, 3: Delayed, 4: Delayed Frozen
    ib.qualifyContracts(contract)
    earn = ib.reqFundamentalData(contract, 'CalendarReport') # Wall Street Horizons Abo notwendig, vieleicht find ich noch andere Quellen
    if type(earn) == str:
        tree = ET.ElementTree(ET.fromstring(earn))
        root = tree.getroot()
        item = tree.find('.//Date')
        data = ib.reqMktData(contract, "106,100", False, False)
        ib.sleep(1)
        now = datetime.now()
        current_time = now.strftime("%H:%M")
        iv = str(data.impliedVolatility)
        pv = str(data.putVolume)
        cv = str(data.callVolume)
        liste = [[item.text, round(float(iv.replace('nan', '0'))*100,2), int(float(pv.replace('nan', '0'))), int(float(cv.replace('nan', '0'))), current_time]]
        df2 = pd.DataFrame(liste)
        df = df.append(df2, ignore_index=True)
        print(ticker, "   ", iv, pv, cv, item.text) # kann auch auskommentiert werden, dient nur zur Kontrolle
        ib.cancelMktData(contract)
    else:
        item = ('01/01/2222') # Pseudo-Earningsdatum, falls die TWS keine liefert
        data = ib.reqMktData(contract, "106,100", False, False)
        ib.sleep(1)
        now = datetime.now()
        current_time = now.strftime("%H:%M")
        iv = str(data.impliedVolatility)
        pv = str(data.putVolume)
        cv = str(data.callVolume)
        liste = [[item, round(float(iv.replace('nan', '0'))*100,2), int(float(pv.replace('nan', '0'))), int(float(cv.replace('nan', '0'))), current_time]]
        df2 = pd.DataFrame(liste)
        df = df.append(df2, ignore_index=True)
        print(ticker, "   ", iv, pv, cv, item) # kann auch auskommentiert werden, dient nur zur Kontrolle
        ib.cancelMktData(contract) 

def main():
    try:
        ticker()
    except Exception as e:
        print(e)
        pass 

if __name__ == "__main__":
    for i in range (2,rows+2): 
        main()

wks2.set_dataframe(df,'H2', copy_head=False) # hier wird alles in die Tabelle geschrieben, wenn alle Symbole durchlaufen sind, da Google nur wenig Zugriffe/min erlaubt, daher der Upload in einem Rutsch.
 
