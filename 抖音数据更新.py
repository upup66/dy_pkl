# -*- coding: utf-8 -*-


import os
import pandas as pd
import numpy as np
import win32com.client


#文件打开保存
def refresh(file_path):
    for file in file_path:
        xlapp = win32com.client.DispatchEx("Excel.Application")
       
        wb = xlapp.Workbooks.Open(file)
        wb.Save()
        wb.Close()
        
        xlapp.Quit()
        print(f'{file}保存成功')
              

# 获取文件夹下所有相同类型的文件路径，返回list类型 target_files('C:\\数据资料',('.json','.xlsx'))
def target_files(path, fmt):
    target = []
    for root, dirs, files in os.walk(path):
        for fn in files:
            name, ext = os.path.splitext(fn)
            if ext in fmt:
                target.append(os.path.join(root, fn))
    return sorted(target, key=os.path.getmtime, reverse=False)


# 订单拆分
def erp(data_erp,key):
    
    data_erp[key]=data_erp[key].astype(str).str.split(',')
    data_erp=data_erp.explode(key,ignore_index=True)

    e=data_erp[data_erp[key].str.contains(':')]
    data_erp.loc[data_erp[key].str.contains(':'),key]=e[key].str.split(':',expand=True)[1]

    data_erp.drop_duplicates(subset=[key],keep='first',inplace=True,ignore_index=True)
    
    return data_erp


# 文件名
def read_fxg(file=None):
    print('正在打开: {}'.format(file))
    # 处理文件名，由文件名确定是哪个店铺的订单数据
    shop_id = {
        'nWgUWak': 'ANU阿奴化妆品旗舰店',
        'nyhGRDYg': 'CMN美妆旗舰店',
        'mOFQCKv': '南京文众美妆专营店',
        'qRJEJZUm': '橙魔美妆专营店',
        'VwNdbkd': '华肌专营店',
        'xcQvtHNE': 'ANU阿奴官方旗舰店',
        'YZFmMcgi': '深野官方旗舰店',
        'rmFDDcMM': '芙可皙授权企业店',
        'rGjguQSC':'LOLLY'
    }
    fn = os.path.splitext(file.split('\\')[-1])[0]
    shop_name = shop_id[fn.split('_')[1][32:]]
    return shop_name
    

# 合并退款信息
def col_con(col1,col2):
    if col1==col2:
        res = col1
    elif (str(col1)=='nan') & (str(col2)!='nan'):
        res = col2
    elif (str(col1)!='nan') & (str(col2)=='nan'):
        res = col1  
    else:
        res = np.nan
    return res


#%% 合并更新
# 存量数据更新
dy = pd.read_pickle(r'C:\数据资料\pkl\抖音pkl\fxg_export_202303291101.pkl')

# dy1 = dy[~dy['支付完成时间'].isnull()].copy()

# st = '2022-01-01 00:00:00'
# et = '2022-12-31 23:59:59'

# dy2 = dy1[(dy1['订单提交时间']>=st)&(dy1['订单提交时间']<=et)].copy()

# dy2['实退金额'].fillna(0,inplace=True)
# dy2['金额'] = dy2['实付金额']-dy2['实退金额']
# dy2['金额'].sum()

# 最后更新时间
dy.groupby(by='店铺').agg({'订单提交时间':'max'})


#%%合并更新
# 待更新数据处理
file = target_files(r'C:\数据资料\pkl\抖音订单', ('.csv')) 
d = []
for f in file:
    da = pd.read_csv(f,dtype=str)
    da['店铺'] = read_fxg(f)
    d.append(da)
df = pd.concat(d)

# 拆分优惠
# clean
for col in ['主订单编号', '子订单编号', '选购商品', '商品规格', '商品ID', '商家编码', 
            '支付方式', '鲁班落地页ID', '达人ID', '仓库ID', '仓库名称']:
    if col in df.columns:
        df[col] = df[col].str.strip()

# 处理日期
for col in ['订单提交时间', '订单完成时间', '支付完成时间', '承诺发货时间']:
    if col in df.columns:
        try:
            df[col] = pd.to_datetime(df[col], errors='raise')
        except ValueError as e:
            print(col, repr(e))
            raise

# 处理数字
for col in ['订单应付金额', '商品单价', '商品数量', '运费', '优惠总金额', '商家改价', 
            '支付优惠', '红包抵扣', '手续费']:
    if col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col].str.replace(',', ''), errors='raise')
        except ValueError as e:
            print(col, repr(e))
            raise

# 其他需要处理的字段，以及添加新的字段，这里从 '商家优惠' 字段拆出 '商家优惠名称' 和 '商家优惠金额', 另外添加 '店铺' 字段
df['split'] = df['商家优惠'].str.split('-')
df['商家优惠名称'] = df['split'].map(lambda x: '-'.join(x[:-1]))
df['商家优惠金额'] = df['split'].map(lambda x: x[-1])
df['商家优惠金额'] = pd.to_numeric(df['商家优惠金额'], errors='raise')
df = df.drop(columns=['split','区', '市', '收件人', '收件人手机号', '省', '街道', '详细地址'])
#%%合并更新

# 合并更新
final = pd.concat([dy,df])
final.drop_duplicates(subset=['主订单编号','子订单编号'],keep='first',inplace = True)

# 实付金额 = 应付-支付优惠 4988964832002768612 ANU
final['实付金额'] = final['订单应付金额'] - final['支付优惠']

#%% 售后退款状态处理
########################################################
# 售后单
sh = target_files(r'C:\数据资料\财务—抖音\售后单', ('.xlsx')) 

# 打开保存
refresh(sh)
# 合并
d1 = []
for f in sh:     
    da = pd.read_excel(f,dtype=str)
    print(f)
    d1.append(da)
sh1 = pd.concat(d1)
sh1.drop_duplicates(subset=['售后单号'],inplace=True)
# 金额
for col in sh1.columns[sh1.columns.str.contains("金额")]:
    sh1[col] = pd.to_numeric(sh1[col], errors='raise')

sh1['退支付优惠（元）'] = sh1['退支付优惠（元）'].str.replace('-','',regex=True)
sh1['退支付优惠（元）'] = pd.to_numeric(sh1['退支付优惠（元）'], errors='raise')
sh1['退支付优惠（元）'].fillna(0,inplace=True)

# sh3 = sh1[(sh1['售后状态']!='售后关闭')&(sh1['退款方式']!='无需退款')].copy()  #退换货和售后未完成无退款
#售后状态筛选
tui = sh1[sh1['售后状态'].isin(['同意退款，退款成功'])].copy()

tui1 = tui.groupby(by='商品单号').agg(
                                    退商品金额=('退商品金额（元）','sum'),
                                    退运费金额=('退运费金额（元）','sum'),
                                    退支付优惠=('退支付优惠（元）','sum'),
                                    退税费金额=('退税费金额（元）','sum'))
tui1.reset_index(inplace=True)

# 实退金额 = 退商品金额-退支付优惠
tui1['实退金额'] = tui1['退商品金额'] - tui1['退支付优惠']

tui1.columns = ['子订单编号', '退商品金额', '退运费金额', '退支付优惠', '退税费金额', '实退金额']

#%%售后状态更新
final.loc[final['子订单编号'].isin(tui1['子订单编号']),'售后状态'] = '同意退款，退款成功'
final.loc[final['售后状态'].isin(['退款成功','已全额退款']),'售后状态'] = '同意退款，退款成功'

# 退款商品{售后状态：'同意退款，退款成功',取消原因：'主品订单售后完成，未发货赠品订单取消'}
# 退款金额合并
final1 = final.merge(tui1,how='left',on='子订单编号',validate='1:1')

final1['实退金额'] = final1.apply(lambda row: col_con(row['实退金额_x'],row['实退金额_y']),axis=1)
final1['退商品金额'] = final1.apply(lambda row: col_con(row['退商品金额_x'],row['退商品金额_y']),axis=1)
final1['退运费金额'] = final1.apply(lambda row: col_con(row['退运费金额_x'],row['退运费金额_y']),axis=1)
final1['退支付优惠'] = final1.apply(lambda row: col_con(row['退支付优惠_x'],row['退支付优惠_y']),axis=1)
final1['退税费金额'] = final1.apply(lambda row: col_con(row['退税费金额_x'],row['退支付优惠_y']),axis=1)

final1.drop(['退商品金额_x', '退运费金额_x', '退支付优惠_x','退税费金额_x', '实退金额_x',
             '退商品金额_y', '退运费金额_y', '退支付优惠_y', '退税费金额_y','实退金额_y'],axis=1,inplace=True)


# 补充退款金额->售后状态为退款，退款金额无的情况
final1['实退金额'] = final1.apply(lambda row : row['实付金额'] if (row['售后状态']=='同意退款，退款成功')&(str(row['实退金额'])=='nan') else row['实退金额'],axis=1)

#%%数据合并更新保存
# 保存
ts = pd.Timestamp.now().strftime('%Y%m%d%H%M') # '%Y%m%d%H%M%S'
final1.to_pickle(r'C:\数据资料\pkl\抖音pkl\fxg_export_{}.pkl'.format(ts))


