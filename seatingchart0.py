# -*- coding: utf-8 -*-
"""
Created on Mon Apr 30 20:37:55 2018

@author: Sarah Zeng

INPUT
Sheet 1: 3 columns where 
column 1 is the alpha/letter component of the seat number
column 2 is the numbering start for the row
column 3 is the numbering end for the row
Sheet 2: 1 column of left-handed seat numbers (alphanumeric)
"""

import pandas as pd

room="PETER108"

filename=room+'.xlsx'
df=pd.read_excel(filename, header=None, sheet_name=0)
leftlist=pd.read_excel(filename, header=None, sheet_name=1)
leftlist=leftlist[0].tolist()

#Make list of seats from chart
tags=[]
for i in range(0,len(df)):
    seats=list(range(df.iloc[i,1],df.iloc[i,2]+1))
    prefix=[df.iloc[i,0]]*len(seats)
    tags=tags+[m+str(n) for m,n in zip(prefix,seats)]

#Export for random assignment
s=pd.Series(tags)
lefty=s.isin(leftlist)
dfe=pd.concat([s,lefty], axis=1)
dfe.columns = ['Seat','Lefty']
dfe.to_excel(room+'_tags.xlsx', index=False)