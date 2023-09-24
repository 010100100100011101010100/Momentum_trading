#FLOW-> Importing important modules -> reading stocks -> creating a final_dataframe and batch api calling -> adding data to the dataframe -> finding the one year percentile return and collecting the top 50 stocks -> building the number of shares column ->
#creating a high quality momentum dataframe along with the api call -> calculating the percentiles -> calculating the HQM score -> keeping the top 50 stocks -> inputing the potfolio and calculating the number of shares -> excel manipulation

#importing the stocks
import math as h
import xlsxwriter as x
from scipy import stats
import pandas as pd
import numpy as np
import requests
from secrets import IEX_CLOUD_API_TOKEN
from statistics import mean

#reading the stocks
stocks=pd.read_csv('sp_500_stocks.csv')

#final_dataframe
m=['Ticker','Stock Price','1 Year','Number of Shares to buy']
final_dataframe=pd.DataFrame(columns=m)
#creating the chunk function along with the symbol string

def chunk(l,n):
    for i in range(0,len(l),n):
        yield l[i:i+n]
symbol_group=list(chunk(stocks,100))
sym_string=[]
for i in range(0,len(symbol_group)):
    sym_string.append(','.join(symbol_group[i]))
#batch api call
for s in sym_string:
    batch=f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=stats,quote&symbols={s}&token={IEX_CLOUD_API_TOKEN}'
    data=requests.get(batch).json()
    final_dataframe=final_dataframe.append(pd.Series([
        s,
        data[s]['quote']['latestPrice'],
        data[s]['stats']['year1ChangePercent'],
        'N/A'
    ],index=m),ignore_index=True)
final_dataframe

#only keeping the top 50 stocks
final_dataframe.sort_values('1 year',ascending=False,inplace=True)
final_dataframe.reset_index(inplace=True,drop=True)
final_dataframe



#calculating the number of shares per stock and this is the end of the basic iteration of our project
portfolio=float(input())
position_size=portfolio/len(final_dataframe.index)
for i in range(len(final_dataframe.index)):
    final_dataframe.loc[i,'Number of Shares to buy']=h.floor(position_size/final_dataframe.loc[i,'Stock Price'])


#creating the high quality momentum dataframe
hqm=['Ticker','Stock Price','1 year','1 year percentile','2 year','2 year percentile','3 year','3 year percentile','6 month','6 month percentile','1 month','1 month percentile','hqm score','number of shares to buy']
h_dataframe=pd.DataFrame(columns=hqm)
for s in sym_string:
    data=requests.get(batch).json()
    h_dataframe=h.dataframe.appen(pd.Series([
        s,
        data[s]['quote']['latestPrice'],
        data[s]['stats']['year1ChangePercent'],
        'na',
        data[s]['stats']['year2ChangePercent'],
        'na',
        data[s]['stats']['year3ChangePercent'],
        'na',
        data[s]['stats']['month6ChangePercent'],
        'na',
        data[s]['stats']['month1ChangePercent'],
        'na',
        'na',
        'na'

    ],index=hqm),ignore_index=True)
h_dataframe


#calculating the percentiles by using the stats module and time period iteration
time_per=['1 year','2 year','3 year','6 month','1 month']
for i in range(len(h_dataframe.index)):
    for t in time_per:
        h_dataframe.loc[i,f'{t} percentile']=stats.percentileofscore(h_dataframe[f'{t}'],h_dataframe.loc[i,'f{t}'])/100


#calculating the hqm score by using the mean of the five percentile and keeping the top 50 stocks with highest hqm scores
for i in h_dataframe.index):
    temp=[]
    for t in time_per:
        temp.append(f'{t} percentile')
    h_dataframe.loc[i,'hqm score']=mean(temp)

#inputing the portfolio and calculating number of shares
portfolio=float(input())
position_size=portfolio/len(h_dataframe.index)
for i in h_dataframe.index:
    h_dataframe.loc[i,'number of shares']=h.floor(position_size/h_dataframe.loc[i,'Stock Price'])

#standard excel manipulation
h_dataframe.sort_values('hqm score',ascending=False,inplace=True)
h_dataframe.reset_index(drop=True,inplace=True)

writer=pd.ExcelWriter('momentum trading.xlsx',engine='xlsxwriter')
h_dataframe.to_excel(writer,sheet_name='momentum trading',index=False)
bg='#0a0a23'
font='#ffffff'

string={
    'font_color':font,
    'bg_color':bg,
    'border':1
    
}

percentage={
    'num_format':'0.0%',
    'font_color':font,
    'bg_color':bg,
    'border':1
    
}

integer={
    'num_format':'0',
    'font_color':font,
    'bg_color':bg,
    'border':1
    
}

dollar={
    'num_format':'$0.00',
    'font_color':font,
    'bg_color':bg,
    'border':1
    
}

#creating column format and executing excel manipulation'
col_format={
    'A':['Ticker',string],
    'B':['Stock Price',dollar],
    'C':['1 year',percentage],
    'D':['1 year percentile',percentage],
    'E':['2 year',percentage],
    'F':['2 year percentile',percentage],
    'G':['3 year',percentage],
    'H':['3 year percentile',percentage],
    'I':['6 month',percentage],
    'J':['6 month percentile',percentage],
    'K':['1 month',percentage],
    'L':['1 month percentile',percentage],
    'M':['hqm score',percentage],
    'N':['number of shares to buy',integer]
}
for i in col_format.keys():
    writer['momentum trading'].set_columns(f'{i}:{i}',20,col_format[i][1])
    writer['momentum trading'].write(f'{i}1',20,col_format[i][0],string)


writer.save()



    






               





