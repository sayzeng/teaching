# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 13:42:51 2020

@author: Sarah Zeng
"""

import string
import pandas as pd

letters = list(string.ascii_uppercase[:-5])
numbers = [1,3,4,6,7,9,10,12,13,15,16,18,19,21]

seats = pd.Series(l + str(n) for l in letters for n in numbers)
print('Number of seats:', len(seats))

seats.to_excel('RecGym.xlsx', index=False)
