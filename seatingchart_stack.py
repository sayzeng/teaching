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

Outputs: Excel spreadsheet with titled 'ROOM_shaped.xlsx'
"""

import pandas as pd
from pandas import ExcelWriter
import numpy as np

# Initiate parameters
room = 'Center119'
NUM_COLS = 6
NUM_CHARTS = 4

# Load in seat assignments
dfo = pd.read_excel(room + '_tags.xlsx', sheet_name='Sheet2')
dfo = dfo[['Seat','Last Name', 'First Name']]
dfo.dropna(inplace=True)

def randseats(df):
    """ Randomises seat assignments """
    
    seats=df['Seat']
    names=df[['Last Name','First Name']]
    seats=seats.sample(frac=1).reset_index(drop=True)    
    return pd.concat([seats,names], axis=1)    

def arrange_table(df, colCount):
    """ Randomises seats and arranges into a table with colCount columns """
    
    df = randseats(df)
    rows = df[df['Last Name'].notnull()] # Removes unassigned seats
    
    agg_cols = np.array_split(rows, colCount)
    for col in agg_cols:
        col.reset_index(drop=True, inplace=True)
    return pd.concat(agg_cols, axis=1)

# Set up Excel sheet names for each seating chart
vno = list(map(str, range(1, NUM_CHARTS + 1)))
versions = ['Sheet' + x for x in vno]

# Create Excel spreadsheet with randomised seating charts
writer = ExcelWriter(room + '_shaped.xlsx')
for v in versions:
    df = arrange_table(dfo, NUM_COLS)
    df.to_excel(writer, v, index=False)   
writer.save()

print('Code complete')