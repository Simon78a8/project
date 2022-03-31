# -*- coding: utf-8 -*-
"""
Created on Thu Mar 31 13:14:25 2022

@author: simon
"""

from openpyxl import load_workbook

wb = load_workbook('produceSales.xlsx')
sheet = wb.active

new_prices = [1.19, 3.08, 1.27]


for row in sheet['B6:B8']:
    for index, cell in enumerate(row):
        cell.value = new_prices[index]
        
        
        
    #for cell in row:
        #print(cell.value)
    #print(row)
    
    
wb.save('new_produceSales.xlsx')
    
    