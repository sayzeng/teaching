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

room = "PETER108"

# Input room seating details
filename = room + '.xlsx'
df = pd.read_excel(filename, header=None, sheet_name=0)
leftlist = pd.read_excel(filename, header=None, sheet_name=1)

# Get seats from chart (list to Series)
seats = []
for i in range(len(df)):
    prefix, rowstart, rowend = df.iloc[i, 0:3]
    seats += (prefix + str(n) for n in range(rowstart, rowend + 1))
seats = pd.Series(seats)

# Export for random assignment
lefty = seats.isin(leftlist)

dfe = pd.concat([seats, lefty], axis=1)
dfe.columns = ['Seat', 'Lefty']
dfe.to_excel(room + '_tags.xlsx', index=False)

print('Code complete.')
