# -*- coding: utf-8 -*-
"""
Created on Sat Mar  4 16:16:54 2023
抖音直播活动分析模板
@author: ANU
"""

import pandas as pd


#%%数据准备
# 季度数据
st1 = '2023-02-11 00:00:00'
et1 = '2023-02-12 23:59:59'
# st1 = '2023-01-01 00:00:00'
# et1 = '2023-03-31 23:59:59'

# 直播时间
# st1 = '2023-03-28 18:58:07'
# et1 = '2023-03-29 00:28:18'


# 抖音全量数据
dy = pd.read_pickle(r"C:\数据资料\pkl\抖音pkl\fxg_export_202303291101.pkl")
dyd = dy[~dy['支付完成时间'].isnull()].copy()    #筛掉未支付订单

# 抖音订单信息表
dyid = pd.read_pickle(r"C:\数据资料\pkl\抖音手机索引.pkl")

# 一个订单号保留一个id
dyid1 = dyid.explode('加密ID')
dyid2 = dyid1.drop_duplicates(subset=['线上订单号'],keep='first')
# 一个订单号保留一个手机号
dyid11 = dyid.explode('手机')
dyid21 = dyid11.drop_duplicates(subset=['线上订单号'],keep='first')

# 数据合并
dyd['加密ID'] = dyd['主订单编号'].map(dyid2.set_index('线上订单号')['加密ID'])
dyd['手机'] = dyd['主订单编号'].map(dyid21.set_index('线上订单号')['手机'])

# 合并手机号,手机号补充加密ID为空的数据
dyd['加密ID'] = dyd.apply(lambda row : row['手机'] if pd.isnull(row['加密ID']) else row['加密ID'],axis=1)

#%% 新老客
# 首购时间
sb = dyd.groupby('加密ID').agg(首购时间=('订单提交时间','min'))
dyd['首购时间'] = dyd['加密ID'].map(sb['首购时间'])
# 订单类型
dyd['订单类型'] = dyd.apply(lambda row: '首购订单' if row['订单提交时间']==row['首购时间'] else '回购订单',axis=1)

# 达人昵称归一
name = dyd.groupby('达人ID').agg(达人昵称=('达人昵称','last'))
dyd['达人昵称'] = dyd['达人ID'].map(name['达人昵称'])
dyd['达人昵称'] = dyd.apply(lambda row: row['达人ID'] if pd.isnull(row['达人昵称']) else row['达人昵称'],axis=1)
dyd.loc[dyd['达人昵称']=='0','达人昵称'] = '商品卡'
dyd.loc[dyd['达人昵称']=='0','达人ID'] = '商品卡'

# 首购达人
dr = dyd[(dyd['订单类型']=='首购订单')&(~dyd['加密ID'].isnull())&(~dyd['达人昵称'].isnull())].copy()
gro_dr = dr.groupby('加密ID').agg(首购直播间=('达人ID','first'))

dyd['首购直播间'] = dyd['加密ID'].map(gro_dr['首购直播间']) 

# 时间拆分
act_dy = dyd[(dyd['订单提交时间'] >= st1 ) & (dyd['订单提交时间'] <= et1)].copy()
bf_dy = dyd[ dyd['订单提交时间'] < st1 ].copy()

# 检验是否存在空ID
a = act_dy[act_dy['加密ID'].isnull()].copy()
assert a.empty

# 时间区间新老客
act_dy['新老客'] = '新客'
act_dy.loc[act_dy['加密ID'].isin(bf_dy['加密ID']),'新老客'] = '老客'


# 直播间新老客
act_dy['是否首购直播间下单'] = act_dy.apply(lambda row: '是' if row['首购直播间']==row['达人ID'] else '否',axis=1)

act_dy['直播间新老客'] = "老客"
act_dy.loc[(act_dy['新老客']=='新客')&(act_dy['是否首购直播间下单']=='是'),'直播间新老客'] = '新客'


#%%导出

# act_dy.to_excel(r"C:\数据资料\抖音订单分析\抖音直播分析\2023-03-28\订单数据.xlsx",index=False)

#%%
# 拼接季度数据
act_dy1 = act_dy.copy()

#%%汇总
act_dy2 = pd.concat([act_dy1,act_dy],ignore_index=True)

# 商品名统一
pin = act_dy2.groupby('商家编码').agg(商品名=('选购商品','last'))
act_dy2['商品名称'] = act_dy2['商家编码'].map(pin['商品名'])


act_dy2.to_excel(r"C:\数据资料\抖音订单分析\抖音直播分析\2023Q1\订单数据.xlsx",index=False)


# 2023-03-29 09:49:43
