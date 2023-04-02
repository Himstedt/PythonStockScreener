import datetime as dt
import pandas as pd
from pandas_datareader import data as pdr
import yfinance as yf
from tkinter import Tk
from tkinter.filedialog import askopenfilename

yf.pdr_override() # <== that's all it takes
start=input("Enter a start year: ")
startyear = int(startyearIn)
now = dt.datetime.now()

# root = Tk()
# ftypes = [7(".xlsm","*.xlsx",".xls")]
# ttl = "Title"
# dir1 = 'C:\\'
# filePath = askopenfilename(filetypes = ftypes, initialdir = dir1, title = ttl)
filePath = r"C:\

stocklist = pd.read_excel(filePath)
stocklist = stocklist.head()
# print(stocklist)

exportList = pd.DataFrame(columns=['Stock', "RS_Rating", "50 Days MA", "150 Day Ma", "200 Day Ma", "52 Week Low", "52 Week High"])

for i in stocklist.index:
    stock = str(stocklist["Symbol"][i])
    RS_Rating = stocklist["RS Rating"][i]
    try:
        df = pdr.get_data_yahoo(stock, start, now)

        smaUsed = [50, 150, 200]
        for x in smaUsed:
            sma = x
            df["SMA_" + str(sma)] = round(df.ilov[:4].rolling(window = sma).mean(),2)

        currentClose = df["Adj Close"][-1]
        moving_average_50 = df["SMA_50"][-1]
        moving_average_150 = df["SMA_150"][-1]
        moving_average_200 = df["SMA_200"][-1]
        low_of_52week = min(df["Adj Close"][-260:])
        high_of_52week = max(df["Adj Close"][-260:])

        try:
            moving_average_200_20past = df["SMA_200"][-20]
        except Exception:
            moving_average_200_20past = 0

        print("Checking "+stock+".....")

        # Condition 1: Current Price > 150 SMA and > 200 SMA
        # Condition 2: 150 SMA and > 200 SMA
        # Condition 3: 200 SMA trending up for at least 1 month (ideally 4-5 months)
        # Condition 4: 50 SMA > 150 SMA and 50 SMA > 200 SMA
        # Condition 5: Current Price > 50 SMA
        # Condition 6: Current Price is at least 30% above 53 week low (Many of the best are up 100-300% before coming out of consolidation)
        # Condition 7: Current Price is within 25% of 52 week high
        # Condition 8: IBD RS rating > 70 and the higher the better

        #exportList = exportList.append({'Stock': stock, "RS_Rating": RS_Rating, "50 Day MA": moving_average_50, "150 Day MA": moving average_150, "200 Day MA": moving_average_200, "52 Week Low": low_of_52week, "52 week High": high_of_52week}, ignore_index = True)

    except Exception:
        print("No data on "+stock")

    print(exportList)
