from datetime import datetime, time
from time import sleep
import requests
import xlwings as xw
import json
import math

# Owned
__author__     = "Rahul Gedam"
__copyright__  = "Copyright 2022-2035, Read Live Option Chain in Excel."
__credits__    = ["Rahul Gedam"]
__license__    = "Public"
__version__    = "1.0.0"
__maintainer__ = "Rahul Gedam"
__email__      = "rahulgedam@gmail.com"
__status__     = "Development"







def fetTopGainerTopLoser():
# URL to fetch the session cookies
 homepage_url = "https://www.nseindia.com"

# URL to the API endpoint
 api_url = "https://www.nseindia.com/api/equity-stockIndices?index=SECURITIES%20IN%20F%26O"

# Headers for the initial request to the homepage
 headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1"
}

# Start a session to persist cookies
 session = requests.Session()

# Make the initial request to the homepage to get cookies
 response = session.get(homepage_url, headers=headers)

# Check if the initial request was successful
 if response.status_code == 200:
    # Add more detailed headers for the API request
    api_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "en-US,en;q=0.5",
        "Referer": "https://www.nseindia.com/",
        "X-Requested-With": "XMLHttpRequest",
        "Connection": "keep-alive"
    }

    # Make the request to the API endpoint with the same session
    response = session.get(api_url, headers=api_headers)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the JSON data
        data = response.json()
        with open("test2.json", "w") as json_file:
            json.dump(data, json_file, indent=4)    
        # Extracting the 'data' field which contains the list of securities
        securities = data['data']

# Sorting the securities by pChange in descending order for top positive changes
        top_positive = sorted(securities, key=lambda x: x['pChange'], reverse=True)[:5]

# Sorting the securities by pChange in ascending order for top negative changes
        top_negative = sorted(securities, key=lambda x: x['pChange'])[:5]

# Printing the top five positive pChange securities
        print("Top 5 Positive pChange:")
        for sec in top_positive:
          print(f"Symbol: {sec['symbol']}, pChange: {sec['pChange']}")

# Printing the top five negative pChange securities
        print("\nTop 5 Negative pChange:")
        for sec in top_negative:
         print(f"Symbol: {sec['symbol']}, pChange: {sec['pChange']}")    

       
    else:
        print(f"Failed to fetch data. Status code: {response.status_code}, Response: {response.text}")
 else:
    print(f"Failed to fetch homepage. Status code: {response.status_code}")











if __name__ == '__main__':
   success = False

   # This loop runs from 9:15 AM to 3:30 PM till Market hours
   while(1):
        t1 = datetime.now()
        currTime = str(t1.hour)+":"+str(t1.minute)
        # From time 9:15 AM to 9:17 AM fetch data and Initialize Files
        #if(time(9,15)<=datetime.now().time()<=time(9,17,30) and not(success)):
        success = fetTopGainerTopLoser()
        #else:
        #  print("Not 9:15 AM Yet, Market yet to start  ", currTime)   

        isFifMin=1
        #while(time(9,20)<=datetime.now().time()<=time(15,31)):
        while(time(8,00)<=datetime.now().time()<=time(23,59)):
            t1 = datetime.now()
            currTime = str(t1.hour) + ":" + str(t1.minute)
            isFifMin+=1
            isFifteenMin = False
            if isFifMin==3:
                isFifteenMin = True
                isFifMin=0
            fetTopGainerTopLoser()
            t1 = datetime.now()
            while(t1.minute%5 !=0 or t1.second !=0):
                t1 = datetime.now()
                sleep(1)
        sleep(1)