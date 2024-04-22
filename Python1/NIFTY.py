import requests
import pandas as pd
import xlwings as xw
import json 
import time 





sym = "NIFTY"
exp_date = '09-Nov-2023'

file = xw.Book("Data.xlsx")
sh1 = file.sheets("Nifty50")
sh2 = file.sheets("ND")

def oc(sym,exp_date ):
     url = "https://www.nseindia.com/api/option-chain-indices?symbol="+sym
     headers = {"accept-encoding":"gzip, deflate, br",
                "accept-language":"en-US,en;q=0.9",
                "referer": "https://www.nseindia.com/get-quotes/derivatives?symbol="+sym,
                "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) ApplelebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.33"}


     response = requests.get(url, headers = headers).text

     data = json.loads(response)

     exp_list = data['records']['expiryDates']

     ce = {}
     pe = {}
     n = 0
     m = 0

     for i in data['records']['data'] :
         
             try:
                 ce[n] = i['CE']
                 n = n+1
             except:
                 pass
             try:
                 pe[m] = i['PE']
                 m = m+1
             except:
                 pass 


     ce_df = pd.DataFrame.from_dict(ce).transpose()
     ce_df.columns += "_CE" 
     pe_df = pd.DataFrame.from_dict(pe).transpose()
     pe_df.columns += "_PE"        

     df = pd.concat([ce_df, pe_df], axis = 1)

     return exp_list, df

while True:
     try:
          data = oc(sym,exp_date)
          sh1.range("A5").value = data[1]
          sh2.range("A1").options(transpose = True).value = data[0]
          time.sleep(60)
     except:
         print("Retrying")      
         time.sleep(5)


