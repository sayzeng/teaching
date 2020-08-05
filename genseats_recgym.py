# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 13:42:51 2020

@author: Sarah Zeng
"""

import string
import pandas as pd

a=list(string.ascii_uppercase[:-5])
b=[1,3,4,6,7,9,10,12,13,15,16,18,19,21]

seats=[]
for l in a:
    seats.extend([l+str(n) for n in b])

len(seats)

seats=pd.Series(seats)

seats.to_excel('RecGym.xlsx',index=False)