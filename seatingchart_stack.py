# -*- coding: utf-8 -*-
"""
Created on Fri Feb  2 07:02:30 2018
@author: Sarah Zeng

This code will read an Excel spreadsheet with a column of
 seat assignments corresponding to column(s) with
 student names and output an Excel spreadsheet that stacks the
 columns next to each other for single-page display during the exam.
 
You choose:
    (1) spreadsheet with your seat assignments
    (2) number of side-by-side stacks
    (3) number of seating charts to generate

Outputs: Excel spreadsheet titled 'chart_shaped.xlsx'

"""

import pandas as pd
from pandas import ExcelWriter
#from pandas import ExcelFile
import numpy as np

#Change 'chart.xlsx' to the file with your seat assignments in three columns: "Seat", "Last Name", "First Name".
room='RecGym'
seatassignments=room+'_tags.xlsx'
#Choose the number of aggregated columns to display
c=6
#Choose number of seatingcharts to make
n=4

#Run the rest of this
dfo=pd.read_excel(seatassignments, sheet_name='Sheet2')
dfo=dfo[['Seat','Last Name', 'First Name']]
dfo.dropna(inplace=True)

def randseats(df):
    seats=df['Seat']
    names=df[['Last Name','First Name']]
    seats=seats.sample(frac=1).reset_index(drop=True)    
    return pd.concat([seats,names], axis=1)    

def arrangetable(df):
    df=randseats(df)
    
    rows=df[df['Last Name'].notnull()]

    x=len(rows)
    r=int(x/c)+1
    #r=48

    jump=np.arange(0,x,r)
    jump=np.append(jump,x)

    df=pd.DataFrame()    
    for i in range(0,c):
        df=pd.concat([df, rows[jump[i]:jump[i+1]].reset_index(drop=True)], axis=1)
    
    return df

vno=list(map(str, range(1,n+1)))
versions=['Sheet' + x for x in vno]

writer=ExcelWriter(room+'_shaped.xlsx')
for v in versions:
    df=arrangetable(dfo)    
    df.to_excel(writer,v,index=False)   
writer.save()

print('Code complete')