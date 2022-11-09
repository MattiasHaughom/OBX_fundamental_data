#%% Imports
import numpy as np
import pandas as pd
import unittest
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from io import StringIO
import os
import itertools
from nltk.tokenize import treebank
from selenium.common.exceptions import NoSuchElementException
import yfinance as yf
from datetime import timedelta
import datetime as datetime



directory = r"C:\Users\pj20\Desktop\New folder"
os.chdir(directory)
#links = pd.read_excel('links.xlsx')
industry = pd.read_excel('OBStickers.xlsx')
allTickers = pd.read_excel('all_tickers.xlsx')
dct = pd.read_excel('norskDict.xlsx')
pos_list=list(dct['positiv'])
neg_list=list(dct['negativ'][:52])
tokenizer = treebank.TreebankWordTokenizer()

def sentiment(sent):
    senti = 0
    words = [word.lower() for word in tokenizer.tokenize(sent)]
    for word in words:
        if word in pos_list:
            senti += 1
        elif word in neg_list:
            senti -= 1
    return senti


#%% OBS DN scraping
def sjekkSektor(tickerLink):
    try:
        sektor = industry.loc[[str(j) in tickerLink for j in industry['ticker']],'sektor'].reset_index(drop = True).values
        return sektor[0]
    except IndexError:
        return 'usikker'


url = 'https://investor.dn.no/#!/Kurser/Aksjer/'

driver = webdriver.Chrome(r'C:\Users\pj20\Desktop\Online Scraping\chromedriver.exe')
driver.minimize_window()
                    
driver.get(url)

rows = len(driver.find_elements(By.XPATH,'//*[@id="dninvestor-content"]/div[1]/div/div[4]/table/tbody/tr'))
before_XPath = '//*[@id="dninvestor-content"]/div[1]/div/div[4]/table/tbody/tr['
aftertd_XPath_3 = ']/td[2]/a'

links = []
for t_row in range(2, (rows + 1)):
    FinalXPath = before_XPath + str(t_row) + aftertd_XPath_3
    links.append(driver.find_element(By.XPATH,FinalXPath).get_attribute("href"))
driver.quit()


listofDfs1 = []
listofDfs2 = []
for i in links:
    url = i
    
    before_XPath = "//*[@id='ticker-numbers']/div/div[1]/div[1]/table/tbody/tr["
    aftertd_XPath_1 = "]/td[1]"
    aftertd_XPath_2 = "]/td[2]"
    aftertd_XPath_3 = "]/td[3]"
    aftertd_XPath_4 = "]/td[4]"
    
    
    first = []
    second = []
    third = []
    fourth = []
    shorts = []
    shortingVol = []
    sisteTek = []
    alleTek = []
    class WebTableTest(unittest.TestCase):
     
        def setUp(self):
            self.driver = webdriver.Chrome(r'C:\Users\pj20\Desktop\Online Scraping\chromedriver.exe')
            self.driver.minimize_window()
     
     
        def test_get_row_col_info_(self):
                driver = self.driver
                driver.get(url)
                
                sleep(10)
         
                rows = len(driver.find_elements(By.XPATH,"//*[@id='ticker-numbers']/div/div[1]/div[1]/table/tbody/tr"))

    
                for t_row in range(2, (rows + 1)):
                    FinalXPath = before_XPath + str(t_row) + aftertd_XPath_1
                    first.append(driver.find_element(By.XPATH,FinalXPath).text)
     
                for t_row in range(2, (rows + 1)):
                    FinalXPath = before_XPath + str(t_row) + aftertd_XPath_2
                    second.append(driver.find_element(By.XPATH,FinalXPath).text)
            
                for t_row in range(2, (rows + 1)):
                    FinalXPath = before_XPath + str(t_row) + aftertd_XPath_3
                    third.append(driver.find_element(By.XPATH,FinalXPath).text)
    
                for t_row in range(2, (rows + 1)):
                    FinalXPath = before_XPath + str(t_row) + aftertd_XPath_4
                    fourth.append(driver.find_element(By.XPATH,FinalXPath).text)
    
    
                # Shorting volume
                try:
                    short = driver.find_element(By.XPATH,"//*[@id='ticker-shorts']/div/div[1]/div").text
                except NoSuchElementException:  
                    short = str(0)
                
                if len(short) == 44:
                    shortingVol.append(float(short[(-1-3):-1].replace(',', '.')))
                elif len(short) > 50:
                    shortingVol.append(float(short[(-15-4):-15].replace(',', '.')))
                elif short == '0':
                    shortingVol.append(float(short))
                

                # Siste tekniske signaler i str
                driver.find_element(By.XPATH,"//*[@id='dninvestor-content']/div[1]/div/div[4]/div[2]/div/div[2]/div[1]").text
                string = driver.find_element(By.XPATH,"//*[@id='dninvestor-content']/div[1]/div/div[4]/div[2]/div/div[2]/div[1]").text
                sisteTekniske = list(itertools.chain.from_iterable(list(map(lambda x: x.split(','),string.split("\n")))))

                sisteSum = 0
                for number,text in enumerate(sisteTekniske):
                    sisteSum += sentiment(sisteTekniske[number])

                sisteTek.append(sisteSum)


                # Tekniske nivåer
                nivåer = driver.find_element(By.XPATH,"//*[@id='dninvestor-content']/div[1]/div/div[4]/div[2]/div/div[2]/div[3]").text
                nivåliste = list(itertools.chain.from_iterable(list(map(lambda x: x.split(','),nivåer.split("\n")))))

                tekniskSum = 0
                for number,text in enumerate(nivåliste):
                    tekniskSum += sentiment(nivåliste[number])
    
                alleTek.append(tekniskSum)

    
    
        def tearDown(self):
            self.driver.close()
            self.driver.quit()
     
    if __name__ == "__main__":
        unittest.main()


    tmp1 = {'estimates':first, '2020': second,'2021': third, '2022': fourth, 'ticker':i,'sektor':sjekkSektor(i)}
    tmp2 = {'shortP':shortingVol,'sisteTek':sisteTek,'alleTek':alleTek,'ticker':i,'sektor':sjekkSektor(i)}
    df1 = pd.DataFrame(data = tmp1)
    df2 = pd.DataFrame(data = tmp2)
    listofDfs1.append(df1)
    listofDfs2.append(df2)

df1 = pd.concat(listofDfs1)
df2 = pd.concat(listofDfs2)
metricData = df1[np.logical_or(np.logical_or(df1['2022']!='-',df1['2021']!='-'),df1['2020']!='-')]
#df = pd.read_csv('estimates'+'.csv')
df1.reset_index(drop= True, inplace = True)
metricData.reset_index(drop= True, inplace = True)


for year in df1.columns[1:4]:
    for number,i in enumerate(df1[year]):
        try:
            if np.logical_and(type(i)==str,"mrd" in i):
                df1.loc[np.logical_and(df1[year]==i,df1[year].index==number),year] = float(i[:-4].replace(',', '.'))*1000000000
            elif np.logical_and(type(i)==str,"mill" in i):
                df1.loc[np.logical_and(df1[year]==i,df1[year].index==number),year] = float(i[:-5].replace(',', '.'))*1000000
            elif type(i) == str:
                df1.loc[np.logical_and(df1[year]==i,df1[year].index==number),year] = float(i.replace(',', '.'))
        except ValueError:
            df1.loc[np.logical_and(df1[year]==i,df1[year].index==number),year] = 0


for year in metricData.columns[1:4]:
    for number,i in enumerate(metricData[year]):
        try:
            if np.logical_and(type(i)==str,"mrd" in i):
                metricData.loc[np.logical_and(metricData[year]==i,metricData[year].index==number),year] = float(i[:-4].replace(',', '.'))*1000000000
            elif np.logical_and(type(i)==str,"mill" in i):
                metricData.loc[np.logical_and(metricData[year]==i,metricData[year].index==number),year] = float(i[:-5].replace(',', '.'))*1000000
            elif type(i) == str:
                metricData.loc[np.logical_and(metricData[year]==i,metricData[year].index==number),year] = float(i.replace(',', '.'))
        except ValueError:
            metricData.loc[np.logical_and(metricData[year]==i,metricData[year].index==number),year] = 0


#metricData.to_csv('estimates.csv')
#df2.to_csv('technical.csv')




#%% Estimating PEG ratio
df = pd.read_csv('estimates.csv')
# importing data from relevant tickers
obx = pd.read_excel('all_tickers.xlsx')

    
tickers = list(obx['ticker'])

for tick in tickers:
    obx.loc[obx['ticker']== tick,'ticker'] = str(tick) + ".OL"

tickers = list(obx['ticker'])


now = datetime.datetime.now()
yesterday = (now - timedelta(days=1)).strftime('%Y-%m-%d')
n = 365
# # Add 2 minutes to datetime object containing current time
past_time = now - timedelta(days=n)

now = now.strftime('%Y-%m-%d')
# Convert datetime object to string in specific format 
past_time = past_time.strftime('%Y-%m-%d')



data1 = yf.download(  # or pdr.get_data_yahoo(...
        # tickers list or string as well
        tickers = tickers,

        # use "period" instead of start/end
        # valid periods: 1d,5d,1mo,3mo,6mo,1y,2y,5y,10y,ytd,max
        # (optional, default is '1mo')
        start=past_time, 
        end=now,

        # fetch data by interval (including intraday if period < 60 days)
        # valid intervals: 1m,2m,5m,15m,30m,60m,90m,1h,1d,5d,1wk,1mo,3mo
        # (optional, default is '1d')
        interval = "1d",

        # group by ticker (to access via data['SPY'])
        # (optional, default is 'column')
        #group_by = 'ticker',

        # adjust all OHLC automatically
        # (optional, default is False)
        auto_adjust = True,

        # download pre/post regular market hours data
        # (optional, default is False)
        prepost = True,

        # use threads for mass downloading? (True/False/Integer)
        # (optional, default is True)
        threads = True,

        # proxy URL scheme use use when downloading?
        # (optional, default is None)
        proxy = None)


yahoo_data = data1.stack()
yahoo_data.reset_index(drop=False, inplace=True)
yahoo_data.rename(columns=({'Date':'date','level_1':'ticker'}), inplace = True)
prices = data1['Close']
prices.index = pd.to_datetime(prices.index)




for number in range(len(df['ticker'])):
    df.loc[number,'ticker'] = df.loc[number,'ticker'].split('/')[6].strip()


def epsRatio(inp,ticker):
    eps = inp.loc[np.logical_and(inp['estimates'] == 'EPS',inp['ticker']==tick),['2020','2021','2022']]
    first = eps['2021'].values[0]
    last = eps['2022'].values[0]
    total = (last - first)/np.abs(first)
    change  = ((total)-1)*100
    return change



def peRatio(inp,ticker):
    peRat = inp.loc[np.logical_and(inp['estimates'] == 'P/E',inp['ticker']==ticker),['2022']]
    try:
        eps = inp.loc[np.logical_and(inp['estimates'] == 'EPS',inp['ticker']==ticker),['2022']]
    except ValueError:
        pass
    except IndexError:
        pass
    if np.logical_and(peRat.values[0][0] == 0, eps.values[0][0] != 0):
        price = prices.loc[prices.index[len(prices)-1],prices.columns == (ticker + ".OL")]
        peRat = price.values[0]/eps
    elif peRat.values[0][0] == 0:
        d = {'data':0}
        peRat = pd.DataFrame(d,index=[0])
    return peRat.values[0][0]



pegRatio = pd.DataFrame(columns= ('ticker','peg'))
for number,tick in enumerate(df['ticker'].unique()):
    try:
        pe = peRatio(df,tick)
        eps = epsRatio(df,tick)
        pegRatio.at[number,'peg'] = pe/eps
        pegRatio.at[number,'ticker'] = tick
    except IndexError:
        pass
    except ValueError:
        pass





