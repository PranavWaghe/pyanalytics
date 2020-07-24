#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jul 24 21:34:46 2020

@author: pranavwaghe
"""

import pandas as pd
import numpy as np

data = pd.read_csv('denco.csv')

data.head()
type(data)
#%%

data.sort_values(by=['custname']) #sorting by customer name
group = data.groupby('custname').aggregate({'revenue': [np.sum,'count']})# summing repeated customers

sh1 = group.nlargest(5,[('revenue', 'count')])# top 5 customers count

#%%
sh2 = group.sort_values([('revenue','sum')], ascending=False)#sort by revenue descending

#%%
group_part_1 = data.groupby('partnum').aggregate({'revenue':[np.sum],'margin':[np.sum]})#grouping by partnumber for revenue
group_part_1.columns# checking column names
p1=group_part_1.sort_values([('revenue', 'sum')],ascending=False) # sorted by part number revenue
p1.head(5)# top 5 revenue

#%%
group_part_2 = data.groupby('partnum').aggregate({'margin':[np.sum]})#grouping by partnumber for margin
p2=group_part_2.sort_values([( 'margin', 'sum')],ascending=False)# sorting data
p2.head(5) # top 5  part by margin
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
group.to_excel(writer, sheet_name='group by customer')
sh1.to_excel(writer, sheet_name='top 5 cust by count')
p1.to_excel(writer, sheet_name='Revenue by part')
p2.to_excel(writer, sheet_name='Margin by part')
writer.save()
