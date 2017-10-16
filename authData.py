# -*- coding: utf-8 -*-
"""
Created on Fri Oct 13 17:34:46 2017

@author: Maddox.Meng
"""

'''req4：审定数据'''

import pandas as pd
import numpy as np
import datetime

wp = r'.\N_workflow\重庆市沙坪坝融兴村镇银行有限责任公司-N-20151231-WP.xlsx'.decode('utf8')
wp2 = r'.\N_workflow\重庆市沙坪坝融兴村镇银行有限责任公司-N-20161231-WP.xlsx'.decode('utf8')

#N_sheets = ['N100', 'N200', 'N300', 'N400']
#
#df = pd.read_excel(file, sheetname='N100', headers = None)
#
#for x in range(0, df.shape[0]):
#    for y in range(0, df.shape[1]):
#        #if df.iloc[x,y] == '审定数据':
#        if df.iloc[x,y] is not np.nan:
#            print x, y, df.iloc[x,y]
#            
#
#def dateToStr(date):
#    return  date.strftime('%Y-%m-%d')



#df = pd.read_excel(wp, sheetname='N100', headers=None)
#
##找到2015年的cell的col，组成list
#yearCols = []
#for x in range(0, df.shape[0]):
#    for y in range(0, df.shape[1]):
#        if isinstance(df.iloc[x,y], pd.Timestamp):
#            #print type(df.iloc[x,y].date()) #'datetime.date'
#            if df.iloc[x,y].year == 2015:
#                yearCols.append(y)
#
#
##找到‘审定数据’，如果col在yearCols里，就记录
#for x in range(0, df.shape[0]):
#    for y in range(0, df.shape[1]):
#        if unicode(df.iloc[x,y]) == u'审定数据' and y in yearCols:
#        #UnicodeWarning: Unicode equal comparison failed to convert both arguments 
#        #to Unicode - interpreting them as being unequal #solve: unicode(df.iloc[x,y])    
#            print x, y, df.iloc[x,y]
#            data_y = y
#            x_min = x + 1
#
##找到‘Check Mapping’，作为x_max
#for x in range(0, df.shape[0]):
#    for y in range(0, df.shape[1]):
#        if unicode(df.iloc[x,y]) == u'Check Mapping':
#            print x, y, df.iloc[x,y]
#            x_max = x - 1
#
#            
##找到2、3列中，审定数据下一行，一直到‘合计’那行的区域里，不为空的cell
#for x in range(x_min, x_max+1):
#    for y in [1,2]:
#        if unicode(df.iloc[x,y]) == u'合计':
#            print x, y, df.iloc[x,y]
#            
#            subj_y = y
#            subj_x_max = x
#            
##在subj_y列中搜审定数据下一行，一直到‘合计’那行的区域里，不为空的cell
#dict = {} 
#k=0           
#y = subj_y            
#for x in range(x_min, x_max+1):
#    '''这里判断需要更多实例确保覆盖所有情况'''
#    #if isinstance(df.iloc[x,y], [np.float, np.nan]) is False:
#    if isinstance(df.iloc[x,y], unicode):    
#        print x, y, df.iloc[x,y]
#        
#        dict[k] = {'subj':df.iloc[x,y], 'value': df.iloc[x,data_y]}
#        k+=1            

'''==============================读被写入文件================================'''
dict=(
{0: {'subj': u'活期存款：', 'value': np.nan},
 1: {'subj': u'  公司客户', 'value': 276688716.5499999},
 2: {'subj': u'  个人客户', 'value': 48668341.94000019},
 3: {'subj': u'小计', 'value': 325357058.49000007},
 4: {'subj': u'定期存款：', 'value': np.nan},
 5: {'subj': u'  公司客户', 'value': 515857443.77},
 6: {'subj': u'  个人客户', 'value': 128324039.64000002},
 7: {'subj': u'小计', 'value': 644181483.41},
 8: {'subj': u'保证金存款', 'value': 61371289.65000004},
 9: {'subj': u'应解汇款', 'value': 590000},
 10: {'subj': u'合计', 'value': 1031499831.5500002}})

from openpyxl import load_workbook
wb = load_workbook(wp2, data_only = True)
ws = wb.get_sheet_by_name('N100')

yearCols2 = []
for x in range(1, ws.max_row+1):
    for y in range(0, ws.max_column):
        if isinstance(ws[x][y].value, datetime.datetime) and ws[x][y].value.year==2015:
            #print x, y, ws[x][y].value.year
            yearCols2.append(y)
            break
        
for x in range(1, ws.max_row+1):
    for y in range(0, ws.max_column):
        if ws[x][y].value == u'审定数据' and y in yearCols2:
            print x, y, ws[x][y].value
            
            data_y = y
            x_min = x + 1
            
for x in range(1, ws.max_row+1):
    for y in range(0, ws.max_column):
        if ws[x][y].value == u'Check Mapping':
            print x, y, ws[x][y].value
            
            x_max = x - 1

for x in range(x_min, x_max + 1):
    for y in range(0, ws.max_column):
        if ws[x][y].value == u'合计':
            print x, y, ws[x][y].value
            
            subj_y = y
            subj_x_max = x            

print '-----------------------'
list = []
y = subj_y
for x in range(x_min, x_max+1):
    #if isinstance(ws[x][y].value, unicode):
    #print x, y, ws[x][y].value
    list.append(ws[x][y].value)

'''有dict，有目标cell位置了，怎么写入'''


'''解决subj非unique的问题'''
i=0        
for item in list:
    if item is not None and list.count(item) > 1:
        #print item,list.count(item), i, list[i-1]
        
        for k in range(1, i+1):
            #print i, k, i-k
            if list.count(list[i-k]) == 1:
                print list[i-k]+list[i],i-k
                break
                
    i+=1  



            