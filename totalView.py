import pandas as pd
from pytrends.request import TrendReq
import datetime
from iso3166 import countries
import dart_fss as dart_fss
import os
import FinanceDataReader as fdr
from pandas import read_csv

#static declare
startdate = '2021-07-19'
API_KEY = 'cede13dfdcf63755563e4a9d4992a32b7123009d'
mycorplist1 = ["삼성전자", "SK하이닉스", "LG전자", "LG유플러스", "한국금융지주", "광주신세계"]
#삼성전자, sk하이닉스, LG전자
def interest_processing(geo, corp_name):
    #pytrend.build_payload(kw_list= keyword, timeframe = timerange, geo = geo, cat = 3)
    key = []
    key.append(corp_name)
    #pytrend.build_payload(kw_list=keyword, timeframe=timerange)
    pytrend.build_payload(kw_list=key, timeframe=timerange)
    dt = pytrend.interest_over_time()
    try:
        del dt['isPartial']
        dt.rename(columns = {'삼성전자' : corp_name}, inplace = True)
        return dt
    except:
        pass

# set date range
today = datetime.datetime.now()
enddate = str(today.date())
timerange = startdate + ' ' + enddate

#keyword choice
print('--------------Search Dart-Fss----------------')
api_key = API_KEY
dart_fss.set_api_key(api_key=api_key)
corp_list = dart_fss.get_corp_list()
corp_list.corps
all = dart_fss.api.filings.get_corp_code()
df = pd.DataFrame(all)
datalist = df[df['stock_code'].notnull()]
mycorplist = [];
for i in range(len(datalist)):
    row = {'index': i, 'stock_code': datalist.iloc[i]['stock_code'], 'corp_name': datalist.iloc[i]['corp_name']}
    mycorplist.append(row)
corpCode = pd.DataFrame(mycorplist)
corpCode['stock_code'].astype(str)
corpCode.to_csv('corpCode.csv')
print('All is Good!')



# total data
keyword = ['삼성전자']
pytrend = TrendReq()
result = pd.DataFrame()
pytrend.build_payload(kw_list= keyword, timeframe= timerange)
total = pytrend.interest_over_time()
period = total.index.values
total = total['삼성전자'].values
result['period'] = period
result['max'] = total

mylist = ['삼성전자']
# data per country
#print('--------------Search TOP 10 GDP Country----------------')
#print('--------------Search Country----------------')
print('--------------Search Corp----------------')
mycountries = [410]
cnt = 0
#for myCorpName in mycorplist:


for i in range(len(mycorplist1)):
#for i in range(len(mylist)):
    #cnt = cnt +1
    #if(cnt == 20):
    #        break
    #myCorpName = mycorplist[len(mycorplist)-1-i]
    myCorpName = mycorplist1[i]
    c = countries[mycountries[0]]
    geo = c.alpha2
    corp_name = mylist[0]
    new = interest_processing(geo, myCorpName)
    print(myCorpName)
    try:
        new.reset_index(drop = True, inplace = True)
        result.reset_index(drop = True, inplace = True)
        result = pd.concat([result, new], ignore_index= False, axis = 1)
    except:
           pass

for k in range(len(result)):
    maxnum = 0
    maxindex = 0
    for i, num in enumerate(result.iloc[k]):
        if i == 0 or i == 1:
            continue
        if num >= maxnum:
            maxnum = num
            maxindex = i - 2
    result.loc[k, 'max'] = result.columns[maxindex+2]

print(result)

file_path = "./data.xlsx"
if os.path.exists(file_path):
    os.remove(file_path)

result.to_excel('data.xlsx')

df_2 = read_csv('corpCode.csv', converters={'stock_code': lambda x: str(x)})
#rint(df_2)
for i in range(len(df_2)):
    if(df_2.loc[i, 'corp_name'] == mycorplist1[0]):
        break

#print(df_2.loc[i, 'stock_code'])
file_path = "./data2.xlsx"
df_3 = fdr.DataReader(df_2.loc[i, 'stock_code'])
if os.path.exists(file_path):
    os.remove(file_path)

df_3.to_excel('data2.xlsx')
print('Good')
