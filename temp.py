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

import numpy as np

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
list_new = []
i=0        
for item in list:
    
    if item is not None and list.count(item) == 1:
        list_new.append(item)
    
    if item is not None and list.count(item) > 1:
        #print item,list.count(item), i, list[i-1]
        
        for k in range(1, i+1):
            #print i, k, i-k
            if list.count(list[i-k]) == 1:
                #print list[i-k]
                print i,k
                print list[i-k]+list[i].strip()#,i-k
                list_new.append(list[i-k]+list[i].strip())
                break
            else:
                pass#list_new.append(item)
                
    i+=1    

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

for k in dict:
    dict[k]['subj'] = list_new[k]
  
