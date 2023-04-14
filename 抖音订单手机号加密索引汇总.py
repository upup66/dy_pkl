# -*- coding: utf-8 -*-
"""
Created on Sat Nov 19 13:52:19 2022
2021年之后的数据--erp
@author: ANU
"""

import pandas as pd
import numpy as np


def erp_explode(data_erp,key):
    
    data_erp[key]=data_erp[key].astype(str).str.split(',')
    data_erp=data_erp.explode(key,ignore_index=True)

    e=data_erp[data_erp[key].str.contains(':')]
    data_erp.loc[data_erp[key].str.contains(':'),key]=e[key].str.split(':',expand=True)[1]
    
    return data_erp


# 手机ID拆分
def upcon(x,y):
    if str(x)==str(y):
        re = x
    elif str(x)=='nan' and str(y)!='nan':
        re = y
    elif str(y)=='nan' and str(x)!='nan' and len(str(x)) < 48:
        re = x            
    else:
        if len(str(x)) < 48:
            re = list(set([x,y]))
        else:
            re = y
    return re

# 合并
def coc(x):
    a = list(x.unique())
        
    if np.nan in a:
        a.remove(np.nan)   
        
    if not a:
        return np.nan
    else:
        if len(a)==1:
            return a[0]
        else:
            return a

        

#%%数据准备
# 抖音聚水潭导出订单
erp_dy = pd.read_excel(r"C:\数据资料\pkl\聚水潭抖音订单导出\订单_2023-03-29_10-46-20.10832679.12283452_1.xlsx",dtype=str)
# erp_dy['下单时间'] = pd.to_datetime(erp_dy['下单时间'])
erp_dy1 = erp_dy[['线上订单号','手机']].copy()
erp_dy2 = erp_explode(erp_dy1,'线上订单号')

erp_dy2['线上订单号'] = erp_dy2['线上订单号'].str.replace('A','',regex=False)

erp_dy2.drop_duplicates(inplace=True)

erp_dy2.columns = ['线上订单号', '手机_导出']

# 验证是否有一个订单号多的手机号的情况
erp = erp_dy2.groupby(by='线上订单号').agg(n = ('手机_导出','unique'),nun = ('手机_导出','nunique'))


# 手机号反爬
erp1 = pd.read_pickle(r"C:\数据资料\手机号订单数据\erp_phone.pkl")
dy = erp1[erp1['shop_site']=='头条放心购'].copy()
dy_ph = dy[['so_id','receiver_mobile_decrypted']].copy()
dy_ph.drop_duplicates(inplace=True)
dy_ph.columns=['线上订单号','手机_反爬']

# 验证两个手机号都能搜到一个订单的情况
dyy = dy_ph.groupby(by='线上订单号').agg(n = ('手机_反爬','unique'),nun = ('手机_反爬','nunique'))

# 合并
final = erp_dy2.merge(dy_ph,how='outer',on='线上订单号')


#%% 字段分列   
# 拆分手机和ID字段
# final['ll'] = final.apply(lambda row :len(str(row['手机_导出'])) ,axis=1) 

final['加密ID'] = final.apply(lambda row : row['手机_导出'] if len(str(row['手机_导出']))==48 else np.nan ,axis=1) 

final['手机'] = final.apply(lambda row : upcon(row['手机_导出'],row['手机_反爬']),axis=1) 

final1 = final.explode('手机')
    
# 按订单合并ID数据
final2 = final1.groupby(by='线上订单号').agg(加密ID = ('加密ID',lambda x : coc(x)), 
                                            手机 = ('手机',lambda x : coc(x))
                                            )
final2.reset_index(inplace=True)

# 判断数据是否是列表
final2['ID_FT'] = final2['加密ID'].apply(lambda x: isinstance(x,list))
final2['手机_FT'] = final2['手机'].apply(lambda x: isinstance(x,list))

#%%空值双向填充
# 1v1数据
final3 = final2[(~final2['加密ID'].isnull())&(~final2['手机'].isnull())&
                (final2['ID_FT']==False)&(final2['手机_FT']==False)].copy()

# 手机号->加密ID
pp = final3.groupby(by='手机').agg(加密ID = ('加密ID',lambda x : coc(x)),num = ('加密ID','nunique'))
pp.reset_index(inplace=True)
pp1 = pp[pp['num']==1].copy()

# 分区
da1 = final2[(final2['加密ID'].isnull())&(final2['手机_FT']==False)].copy()
da2 = final2[~final2['线上订单号'].isin(da1['线上订单号'])].copy()

# 匹配
da1['加密ID'] = da1['手机'].map(pp1.set_index('手机')['加密ID'])
# 合并
final_data2 = pd.concat([da1,da2])

##############
# 1v1数据
final4 = final_data2[(~final_data2['加密ID'].isnull())&(~final_data2['手机'].isnull())&
                (final_data2['ID_FT']==False)&(final_data2['手机_FT']==False)].copy()

#加密ID->手机号
dd = final4.groupby(by='加密ID').agg(手机 = ('手机',lambda x : coc(x)),num = ('手机','nunique'))
dd.reset_index(inplace=True)
dd1 = dd[dd['num']==1].copy()
# 分区
da11 = final_data2[(final_data2['手机'].isnull())&(final_data2['ID_FT']==False)].copy()
da21 = final_data2[~final_data2['线上订单号'].isin(da11['线上订单号'])].copy()
# 匹配
da11['手机'] = da11['加密ID'].map(dd1.set_index('加密ID')['手机'])
# 合并
final_data3 = pd.concat([da11,da21])


# 重设index
final_data3.reset_index(drop=True, inplace=True) 


# 判断数据是否是列表
final_data3['ID_FT'] = final_data3['加密ID'].apply(lambda x: isinstance(x,list))
final_data3['手机_FT'] = final_data3['手机'].apply(lambda x: isinstance(x,list))


len(final_data3[final_data3['加密ID'].isnull()])


# # 抖音全量数据
# dy = pd.read_pickle(r"C:\数据资料\pkl\抖音pkl\fxg_export_202301041451.pkl")

# a = dy[~dy['主订单编号'].isin(final_data3['线上订单号'])].copy()
#%%导出
final_data3.to_pickle(r'抖音手机索引.pkl')


