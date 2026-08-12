#!/usr/bin/python
# -*- encoding: utf-8 -*-
# File    :   data_check.py
# Time    :   2026/08/05 08:56:26
# Author  :   Hsu, Liang-Yi 
# Email:   yi75798@gmail.com
# Description :

import pandas as pd
from pandas.api.types import (
    is_numeric_dtype,
    is_string_dtype,
    is_datetime64_any_dtype)
import numpy as np
import pyreadstat
from datetime import datetime

#### data import
# check_date = '0805'
# raw_data_path = r'C:\Users\yi5011\Documents\GitHub\srod_dp\data_s_test.sav'

def data_import(raw_data_path:str):
    '''
    讀取原始資料，支援spss的sav格式、csv格式

    Parameters
    ----------
    raw_data_path: str
        原始資料路徑
    
    Returns
    ---------
    pandas.Dataframe
        讀取後的原始資料pandas dataframe
    '''
    if raw_data_path[-3:] == 'sav':
        df, meta = pyreadstat.read_sav(
            raw_data_path,
            apply_value_formats=False,
            user_missing=True)
        return df
        
    elif raw_data_path[-3:] == 'csv':
        df = pyreadstat.read_csv(
            raw_data_path,
            apply_value_formats=False,
            user_missing=True)
        return df

### 定義遺漏值
def def_na(df, numeric_na=None):
    '''
    定義遺漏值，避免後續判斷是否為遺漏值時出錯。
    注意: 此處的遺漏值指資料本身缺失值，並非問卷資料檔定義的遺漏值(如99)

    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    numeric_na: any type
        欲對數值型資料定義的遺漏值，預設為 None
    
    Returns
    ---------
    pandas.Dataframe
        定義遺漏值後的pandas dataframe
    '''
    for i in df.columns:
        if is_numeric_dtype(df[i]):
            df[i] = df[i].fillna(None)
        elif is_string_dtype(df[i]):
            df[i] = df[i].fillna('')
        elif is_datetime64_any_dtype(df[i]):
            df[i] = df[i].fillna(np.nan)
    return df

#### 必要檢查項目
### 不合理值
## 數值以外題型
def unreason_single(df, result, var:str, values:list,
             check_date=datetime.today().strftime('%m%d'), Id='id', wno='wno', birth='birth2026', gender='gender2026', des=f'出現不合理值或遺漏值'):
    '''
    類別變項不合理值檢核，如不應出現的數字
    注意: 記得放入跳答碼、無反應選項

    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    result: pandas.Dataframe
        要輸出檢核結果的dataframe
    var: str
        變項名稱
    values: list
        合理值
    check_date: str
        檢核日期，預設為檢核程式執行當日
    Id: str
        樣本編號變項名稱，預設為'id'
    wno: str
        會員編號變項名稱，預設為'wno'
    birth: str
        出生年變項名稱，預設為'birth2026'
    gender: str
        性別變項名稱，預設為'gender2026'
    des: str
        要輸出的不符合品說明，預設為'出現不合理值或遺漏值'
    
    Returns
    ---------
    直接輸出檢核結果至result參數指定的dataframe
       
    '''
    for i in df.index:
        if (df.loc[i, var] not in values) or (np.isnan(df.loc[i, var])):
            result.loc[len(result)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       f'{var}={df.loc[i, var]}',
                                       des]

## 數值題
def unreason_con(df, result, var:str, values_mM:list, nr_values=[],
                 check_date=datetime.today().strftime('%m%d'), Id='id', wno='wno', birth='birth2026', gender='gender2026', des=f'出現不合理值或遺漏值'):
    '''
    連續變項不合理值檢核，如超出範圍的數值
    注意: 記得放入跳答碼、無反應選項

    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    result: pandas.Dataframe
        要輸出檢核結果的dataframe
    var: str
        變項名稱
    values_mM: list
        合理值上下限，第一位為下限，第二位為上限
    nr_values: list
        無反應選項及跳答碼清單，預設留空
    check_date: str
        檢核日期，預設為檢核程式執行當日
    Id: str
        樣本編號變項名稱，預設為'id'
    wno: str
        會員編號變項名稱，預設為'wno'
    birth: str
        出生年變項名稱，預設為'birth2026'
    gender: str
        性別變項名稱，預設為'gender2026'
    des: str
        要輸出的不符合品說明，預設為'出現不合理值或遺漏值'
    
    Returns
    ---------
    直接輸出檢核結果至result參數指定的dataframe
        
    '''
    for i in df.index:
        if (((df.loc[i, var] < values_mM[0]) or (df.loc[i, var] > values_mM[1])) & (df.loc[i, var] not in nr_values)) or (np.isnan(df.loc[i, var])):
            result.loc[len(result)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       f'{var}={df.loc[i, var]}',
                                       des]

### 跳續答
## 數值題以外題型
def skip(df, result, var1:str, val1:list, var2:str, val_skip:list,
         check_date=datetime.today().strftime('%m%d'), Id='id', wno='wno', birth='birth2026', gender='gender2026', des='應答未答或不應答卻答'):
    '''
    跳答邏輯檢核，包含「不該答而答」「該答而未答」
    
    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    result: pandas.Dataframe
        要輸出檢核結果的dataframe
    var1: str
        邏輯題變項名稱，僅支援類別變項
    val1: list
        觸發跳答邏輯的值
    var2: str
        被跳答變項名稱
    val_skip:
        被跳答題跳答碼
    check_date: str
        檢核日期，預設為檢核程式執行當日
    Id: str
        樣本編號變項名稱，預設為'id'
    wno: str
        會員編號變項名稱，預設為'wno'
    birth: str
        出生年變項名稱，預設為'birth2026'
    gender: str
        性別變項名稱，預設為'gender2026'
    des: str
        要輸出的不符合品說明，預設為出現不合理值或遺漏值
    
    Returns
    ---------
    直接輸出檢核結果至result參數指定的dataframe
        
    '''
    for i in df.index:
        if (df.loc[i, var1] in val1) and (df.loc[i, var2] in val_skip):
            result.loc[len(result)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       f'{var1}={df.loc[i, var1]};{var2}={df.loc[i, var2]}',
                                       des]
        if (df.loc[i, var1] not in val1) and (df.loc[i, var2] not in val_skip):
            result.loc[len(result)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       f'{var1}={df.loc[i, var1]};{var2}={df.loc[i, var2]}',
                                       des]

### 複選題至少選1個
def multi_1(df, result, var:list,
            check_date=datetime.today().strftime('%m%d'), Id='id', wno='wno', birth='birth2026', gender='gender2026', des='複選題至少應選1個'):
    '''
    複選題數量檢核

    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    result: pandas.Dataframe
        要輸出檢核結果的dataframe
    var: list
        納入計算的複選題變項清單
    check_date: str
        檢核日期，預設為檢核程式執行當日
    Id: str
        樣本編號變項名稱，預設為'id'
    wno: str
        會員編號變項名稱，預設為'wno'
    birth: str
        出生年變項名稱，預設為'birth2026'
    gender: str
        性別變項名稱，預設為'gender2026'
    des: str
        要輸出的不符合品說明，預設為'複選題至少應選1個'
    
    Returns
    ---------
    直接輸出檢核結果至result參數指定的dataframe
        
    '''
    for i in df.index:
        if df.loc[i, var].sum() == 0:
            var_text = ';'.join([f'{v}={df.loc[i, v]}' for v in var])
            result.loc[len(result)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       var_text,
                                       des]

### 複選題無反應不可與其他一起選
def multi_nr(df, result, var:list, nr:list,
             check_date=datetime.today().strftime('%m%d'), Id='id', wno='wno', birth='birth2026', gender='gender2026', des='複選題無反應選項不應與其他一起出現'):
    '''
    複選題無反應選項邏輯檢核，無反應選項不應與其他選項一起出現
    
    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    result: pandas.Dataframe
        要輸出檢核結果的dataframe
    var: str
        變項名稱
    nr: list
        無反應選項變項名稱
    check_date: str
        檢核日期，預設為檢核程式執行當日
    Id: str
        樣本編號變項名稱，預設為'id'
    wno: str
        會員編號變項名稱，預設為'wno'
    birth: str
        出生年變項名稱，預設為'birth2026'
    gender: str
        性別變項名稱，預設為'gender2026'
    des: str
        要輸出的不符合品說明，預設為'複選題無反應選項不應與其他一起出現'
    
    Returns
    ---------
    直接輸出檢核結果至result參數指定的dataframe
        
    '''
    for i in df.index:
        if df.loc[i, var].sum() != 0 and df.loc[i, nr].sum() != 0:
            var_text = ';'.join([f'{v}={df.loc[i, v]}' for v in var+nr])
            result.loc[len(result)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       var_text,
                                       des]

### OE檢核
def oe_list(df, oe_answer, var:str,
            check_date=datetime.today().strftime('%m%d'), Id='id', wno='wno', birth='birth2026', gender='gender2026', des='開放題確認'):
    '''
    列出開放題答案
    
    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    oe_answer: pandas.Dataframe
        要輸出開放題答案的dataframe
    var: str
        變項名稱
    check_date: str
        檢核日期，預設為檢核程式執行當日
    Id: str
        樣本編號變項名稱，預設為'id'
    wno: str
        會員編號變項名稱，預設為'wno'
    birth: str
        出生年變項名稱，預設為'birth2026'
    gender: str
        性別變項名稱，預設為'gender2026'
    des: str
        要輸出的不符合品說明，預設為'開放題確認'
    
    Returns
    ---------
    直接輸出檢核結果至result參數指定的dataframe
        
    '''
    for i in df.index:
        if df.loc[i, var] != '':
            oe_answer.loc[len(oe_answer)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       f'{var}={df.loc[i, var]}',
                                       des]

### OE 不正常符號確認
def oe_abnormal(df, oe_answer, var:str,
               symbols=['*', '{', '}', '[', ']', '&'],
               check_date=datetime.today().strftime('%m%d'), Id='id', wno='wno', birth='birth2026', gender='gender2026', des='開放題不正常符號確認'):
    '''
    開放題不正常符號檢核
    
    Parameters
    ----------
    df: pandas.Dataframe
        原始資料
    oe_answer: pandas.Dataframe
        要輸出檢核結果的dataframe
    var: str
        變項名稱
    symbols: list
        不正常符號清單，預設為'['*', '{', '}', '[', ']', '&']'
    check_date: str
        檢核日期，預設為檢核程式執行當日
    Id: str
        樣本編號變項名稱，預設為'id'
    wno: str
        會員編號變項名稱，預設為'wno'
    birth: str
        出生年變項名稱，預設為'birth2026'
    gender: str
        性別變項名稱，預設為'gender2026'
    des: str
        要輸出的不符合品說明，預設為'開放題不正常符號確認'
    
    Returns
    ---------
    直接輸出檢核結果至result參數指定的dataframe
        
    '''
    for i in df.index:
        if df.loc[i, var] != '':
            for s in symbols:
                if s in df.loc[i, var]:
                    oe_answer.loc[len(oe_answer)] = [check_date,
                                            df.loc[i, Id],
                                            df.loc[i, wno],
                                            df.loc[i, gender],
                                            df.loc[i, birth],
                                            f'{var}={df.loc[i, var]}',
                                            des]
                    
#### test
if __name__ == '__main__':
    ## 定義路徑
    # 原始資料檔
    raw_data_path = r'C:\Users\yi5011\Documents\GitHub\srod_dp\data_s_test.sav'
    # 輸出檔路徑
    output_path = r'C:\Users\yi5011\Documents\GitHub\srod_dp'

    # 讀取檔案
    
    df = data_import(raw_data_path)

    # 定義數值遺漏值
    df = def_na(df)

    # 建立固定變項
    Id = 'id' # 樣本編號
    wno = 'wno' # 會員編號
    gender = 'gender2026' # 性別
    birth = 'birth2026' # 出生年

    # 建立輸出檔
    # 不符合品
    result = pd.DataFrame(columns=['檢核日期', '樣本編號', 'wno', '性別', '年次', '變項名稱', '不符合說明'])
    # 開放題清單
    oe_answer = pd.DataFrame(columns=['檢核日期', '樣本編號', 'wno', '性別', '年次', '變項名稱', '不符合說明'])

    # 連續不合理
    unreason_con(df, result, 'vA1', [0, 9999999], [])

    # 跳答
    skip(df, result, 'vB1', [1], 'vB1U', [9999999996])
    skip(df, result, 'vB1', [2], 'vB1D', [9999999996])

    # 複選數量檢核
    F1 = ['vF1m'+str(n) for n in range(1, 12)]
    multi_1(df, result, F1)

    # 開放題列出
    oe_list(df, oe_answer, 'vC3o7')

    # 開放題不正常符號列出
    oe_abnormal(df, oe_answer, 'vC3o7')


    







            


