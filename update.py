# -*- coding: utf-8 -*-
"""
Created on Thu Mar 31 13:14:25 2022

@author: simon
"""

from openpyxl import load_workbook  ##import workbook

wb = load_workbook('produceSales.xlsx')  #load workbook from file
sheet = wb.active                        #activate sheet

new_prices = [1.19, 3.08, 1.27]


for row in sheet['B6:B8']:                            ##cells with prices to be updated
    for index, cell in enumerate(row):
        cell.value = new_prices[index]
        
        
        
    
    
    
wb.save('new_produceSales.xlsx')
    
    
