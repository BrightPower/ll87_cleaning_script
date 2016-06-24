# -*- coding: utf-8 -*-
"""

LL87 Data Cleaning Script

Created on Sat Nov  8 10:14:23 2014

@author: conor
"""
import pandas as pd

filepath = ("LL87 2013+2014 Energy Audit Data.xlsx")
column_list="A,C,G,H,J,K,L,M,N,P,U,AG,AO:AV,BD:BT,CO:CV,CY,DE,FL,FQ,GK,GR,IN,IT,UX:AIT,AJH:AJS,ANB"
ll87data = pd.io.excel.read_excel(filepath,sheetname="All (1)s",parse_cols=column_list,skiprows=3)

### REMOVING DUPLICATES
#remove perfect dupes
print("Look for perfect dupes")
print(ll87data.duplicated().value_counts())
ll87data = ll87data[~ll87data.duplicated()]
print("After removing perfect dupes")
print(ll87data.duplicated().value_counts())
#remove BBL and BIN dupes
print("Look for dupes just of BBL and BIN")
print(ll87data.duplicated(subset=["BBL","BIN"],take_last=True).value_counts())
ll87data = ll87data[~ll87data.duplicated(subset=["BBL","BIN"],take_last=True)]


### FIXING BOROUGH NAMES USING REGULAR EXPRESSIONS
bordict={r"(?i)bronx":"Bronx",
         r"(?i)manhattan":"Manhattan"
         }
ll87data["borough"] = ll87data["borough"].replace(bordict,regex=True)

### Handle poorly entered square feet and convert to numbers
ll87data["Gross Floor Area"] = ll87data["Gross Floor Area"].replace(r"[ \t]?sqft\.?","",regex=True)
ll87data["Gross Floor Area"] = ll87data["Gross Floor Area"].replace(r",","",regex=True)
ll87data["Gross Floor Area"] = ll87data["Gross Floor Area"].convert_objects(convert_dates=False, convert_numeric=True)

### Converting many fields to numbers and some cleaning
ll87data["weather normalized source EUI"] = ll87data["weather normalized source EUI"].convert_objects(convert_dates=False, convert_numeric=True)
ll87data.rename(columns={"weather normalized source EUI":"LL87 Source EUI"}, inplace=True)
ll87data["early compliance"] = ll87data["early compliance"].replace(r"(?i)no","No",regex=True)
ll87data["Year of construction/substantial Rehabilitation"] = ll87data["Year of construction/substantial Rehabilitation"].convert_objects(convert_numeric=True)
ll87data["Central Distribution Type"] = ll87data["Central Distribution Type"].replace(r"(?i)2-Pipe Steam","2-pipe Steam",regex=True)
ll87data["Total Estimated.4"] = ll87data["Total Estimated.4"].convert_objects(convert_numeric=True)

ll87data["Measure Name"] = ll87data["Measure Name"].str.rstrip()
ll87data["Category"] = ll87data["Category"].str.rstrip()
ll87data["Implementation Cost"] = ll87data["Implementation Cost"].convert_objects(convert_numeric=True)
ll87data["Total Annual Cost Savings"] = ll87data["Total Annual Cost Savings"].convert_objects(convert_numeric=True)
ll87data["Simple Payback"] = ll87data["Simple Payback"].convert_objects(convert_numeric=True)
ll87data["Total Annual Energy Savings"] = ll87data["Total Annual Energy Savings"].convert_objects(convert_numeric=True)

# All use of rstrip below and above here are to remove trailing whitespaces
for i in range(24):
    ll87data["Category."+ str(i+1)] = ll87data["Category."+ str(i+1)].str.rstrip()
    ll87data["Measure Name."+ str(i+1)] = ll87data["Measure Name."+ str(i+1)].str.rstrip()
    ll87data["Implementation Cost."+ str(i+1)] = ll87data["Implementation Cost."+ str(i+1)].convert_objects(convert_numeric=True)
    ll87data["Total Annual Cost Savings."+ str(i+1)] = ll87data["Total Annual Cost Savings."+ str(i+1)].convert_objects(convert_numeric=True)
    ll87data["Simple Payback."+ str(i+1)] = ll87data["Simple Payback."+ str(i+1)].convert_objects(convert_numeric=True)
    ll87data["Total Annual Energy Savings."+ str(i+1)] = ll87data["Total Annual Energy Savings."+ str(i+1)].convert_objects(convert_numeric=True)

### Fixing number of floors using Regular Expressions
floors_dict={r"(?i)two":"2",
             r"(?i)five":"5",
             r"(?i)5/4":"5",
             r"(?i)six":"6",
             r"(?i)5/6":"6",
             r"(?i)seven":"7",
             r"(?i)7 for.*":"7",
             r"(?i)fifteen":"15",
             r"(?i)24.*Pent.*":"25",
             r"(?i)24.*Pent.*":"25",
             r"(?i)31.*Bard.*":"31",
             r"(?i)35 for.*":"35",
             r"(?i)22,.*":"22",
             r"(?i)2-3":"3",
             r"(?i)1;.*":"3",
             r"(?i)12;.*":"12",
             r"(?i)114 West.*":"12"
}
ll87data["# of above grade floors"] = ll87data["# of above grade floors"].replace(floors_dict,regex=True)
ll87data["# of above grade floors"] = ll87data["# of above grade floors"].convert_objects(convert_numeric=True)

with pd.ExcelWriter('cleanll87data_2013_2014.xlsx') as writer:
    ll87data.to_excel(writer,'Raw Data',index=False)
  
    