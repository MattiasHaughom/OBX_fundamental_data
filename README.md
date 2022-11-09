# OBX_fundamental_data
Scraping stock news and metrics from the Norwegian newspaper Dagens Næringsliv's ("DN") Investor page.


## Data 
Online freely available fundamental data (EBITA, sales, P/E etc.) are not available to download for most Norwegian companies. The various sources that do present some data are often in conflict with other sources, motivating me to retreive data from DN as this source is viewed as more reliable. Note that I used chromedriver to scrape the newspaper, and that scraping the fundamental data will take at least 1.5 hours to run because it's manouvering into around 300+ companies' pages.

The other data I have uploaded into the ".zip" folder:  
all_tickers.xlsx - list of all tickers and their company name in the OBX stock index  
norskDict.xlsx - list of positve and negative words for norwegian stock news, for sentiment analysis  
OBStickers.xlsx - list of tickers and their sector/industry, where these are available  

## Script
Takes around 1.5 hours to run, and could possibly be improved by only getting the companies where new data has been released recently. The current code retreives a list of all the links present in the https://investor.dn.no/#!/Kurser/Aksjer/ overview in order to get the most updated list of stocks. It subsequently downloads the "Estimates" table for each stock where it is present (about half the companies), the latest technical indicators, the secton called "Tekniske nivåer" which is a summary of different types of technical analysis and lastly the shorting percentage (% of shorted stocks).

After the data has been downloaded I aggrigate the textual technical analysis data into one dataframe called df2, which compares the number of positive and negative words used to describe the technical data. In addition to this I transform the fundamental data into real numbers and remove the companies where fundamental data was not available.

The last section calculates the PEG ratio for companies where P/E ratio and EPS for the last two years is available.
