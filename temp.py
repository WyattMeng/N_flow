# -*- coding: utf-8 -*-
"""
Created on Tue Oct 10 14:35:23 2017

@author: Maddox.Meng
"""

'''temp
Requirement3:A3的WP到N的WP
Description:
A3-WP的Sheet.TB的K列 
写到N-WP的Sheet.N100,Sheet.N200,Sheet.N300,Sheet.N400的checking Mapping区域的TB列,条件是科目码对应.
'''

#import openpyxl
#import pandas as pd
#print openpyxl.__version__
#print pd.__version__ #0.20.1

A3_WP = r'C:\Workspace\AuditAutomation_N\N_workflow\重庆市沙坪坝融兴村镇银行有限责任公司-A3-20161231-WP.xlsx'.decode('utf-8')  
N_WP  = r'C:\Workspace\AuditAutomation_N\N_workflow\重庆市沙坪坝融兴村镇银行有限责任公司-N-20161231-WP-light.xlsx'.decode('utf-8')

list = [None,
 u'活期存款：',
 u'  公司客户',
 u'  个人客户',
 u'小计',
 None,
 u'定期存款：',
 u'  公司客户',
 u'  个人客户',
 u'小计',
 None,
 u'保证金存款',
 u'应解汇款',
 None,
 u'合计',
 None,
 None,
 None,
 None]

#for i in list:
#    if i is not None:
#        print i, list.index(i)
i=0        
for item in list:
    if item is not None and list.count(item) > 1:
        #print item,list.count(item), i, list[i-1]
        
        for k in range(1, i+1):
            #print i, k, i-k
            if list.count(list[i-k]) == 1:
                #print list[i-k]
                print list[i-k]+list[i],i-k
                break
                
    i+=1    

