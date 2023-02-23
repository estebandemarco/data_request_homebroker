# -*- coding: utf-8 -*-
"""
Created on Sun Feb 12 12:57:22 2023

@author: urbin
"""
from pyhomebroker import HomeBroker
import xlwings as xw
import time
import pandas as pd


#User and broker validation

broker = "Broker Number"
dni = "XXXX"
user = "XXXXXX"
password = "XXXXX"


# API variables
hb = HomeBroker(int(broker))

hb.auth.login(dni=dni, user=user, password=password, raise_exception=True)

hb.online.connect()

#List of stocks tickers

tickers=["BMA","CEPU","CRES","GLOB","GOOGL","IRSA","MELI","MIRG","MSFT","SUPV","TECO2","TGNO4", "VIST","XLF" ]

def get_intraday_history(ticker):
    #Get the intraday history for the given stock
    x = hb.history.get_intraday_history(ticker)[-1:]
    return x

def get_intraday_history_for_tickers(tickers):
    #Get the intraday history for a list of stocks
    resultado= []
    for ticker in tickers:
        x = get_intraday_history(ticker)
        resultado.append(x)
    return resultado

#Variable to introduce spreadsheet location
wb = xw.Book('D:\Downloads\inversiones.xlsx')

#Variable to introduce excel sheet
sht = wb.sheets['Hoja1']

#Loop for continuos update
while True:
    try:
        dato = pd.concat(get_intraday_history_for_tickers(tickers))
        print(dato)
        print("Online", time.strftime("%H:%M:%S"))
        sht.range('B1').value = dato
        time.sleep(30)
    except KeyboardInterrupt:
        print("Keyboard interrupt received. Exiting...")
        break
















