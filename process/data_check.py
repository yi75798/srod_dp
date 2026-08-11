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

#### data import
check_date = '0805'
raw_data_path = r'C:\Users\yi5011\Documents\GitHub\srod_dp\data_s_test.sav'

### sav
df, meta = pyreadstat.read_sav(
    raw_data_path,
    apply_value_formats=False,
    user_missing=True
)

print(df.head())

### csv

### 定義遺漏值
for i in df.columns:
    if is_numeric_dtype(df[i]):
        print(i, '=', df[i].dtype)
    elif is_string_dtype(df[i]):
        print(i, '=', df[i].dtype)
    elif is_datetime64_any_dtype(df[i]):
        print(i, '=', df[i].dtype)

for i in df.columns:
    if is_numeric_dtype(df[i]):
        df[i] = df[i].fillna(np.nan)
    elif is_string_dtype(df[i]):
        df[i] = df[i].fillna('')
    elif is_datetime64_any_dtype(df[i]):
        df[i] = df[i].fillna(np.nan)

### 固定變項
Id = 'id'
wno = 'wno'
gender = 'gender2026'
birth = 'birth2026'

#### result output
result = pd.DataFrame(columns=['檢核日期', '樣本編號', 'wno', '性別', '年次', '變項名稱', '不符合說明'])
oe_list = pd.DataFrame(columns=['檢核日期', '樣本編號', 'wno', '性別', '年次', '變項名稱', '不符合說明'])

#### 必要檢查項目
### 不合理值
## 數值以外題型
def unreason_single(df, var:str, values:list,
             Id=Id, wno=wno, birth=birth, gender=gender, des=f'出現不合理值或遺漏值'):
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
def unreason_single(df, var:str, values_mM:list,
             Id=Id, wno=wno, birth=birth, gender=gender, des=f'出現不合理值或遺漏值'):
    for i in df.index:
        if (not values_mM[0] <= df.loc[i, var] <= values_mM[1]) or (np.isnan(df.loc[i, var])):
            result.loc[len(result)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       f'{var}={df.loc[i, var]}',
                                       des]

### 跳續答
## 數值題以外題型
def skip(df, var1:str, val1:list, var2:str, val_skip:list,
          Id=Id, wno=wno, birth=birth, gender=gender, des='應答未答或不應答卻答'):
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
def multi_1(df, var:list,
            Id=Id, wno=wno, birth=birth, gender=gender, des='複選題至少應選1個'):
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
def multi_nr(df, var:list, nr:list,
             Id=Id, wno=wno, birth=birth, gender=gender, des='複選題無反應選項不應與其他一起出現'):
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
def oe_listed(df, var:str,
            Id=Id, wno=wno, birth=birth, gender=gender, des='開放題確認'):
    for i in df.index:
        if df.loc[i, var] != '':
            oe_list.loc[len(oe_list)] = [check_date,
                                       df.loc[i, Id],
                                       df.loc[i, wno],
                                       df.loc[i, gender],
                                       df.loc[i, birth],
                                       f'{var}={df.loc[i, var]}',
                                       des]

### OE 不正常符號確認
def oe_cintent(df, var:str,
               symbols=['*', '{', '}', '[', ']', '&'],
               Id=Id, wno=wno, birth=birth, gender=gender, des='開放題不正常符號確認'):
    for i in df.index:
        if df.loc[i, var] != '':
            for s in symbols:
                if s in df.loc[i, var]:
                    oe_list.loc[len(result)] = [check_date,
                                            df.loc[i, Id],
                                            df.loc[i, wno],
                                            df.loc[i, gender],
                                            df.loc[i, birth],
                                            f'{var}={df.loc[i, var]}',
                                            des]
                    
#### test
if __name__ == '__main__':
    df = pd.DataFrame({'id':[1,2,3,4,5],
                       'wno':[1,2,3,4,5],
                       'gender2026':['M','F','M','F','X'],
                       'birth2026':[2000,2001,2002,2003,2004],
                       'v1':[1, 2, 5, 7, 99],
                       'v2_1':[0, 1, 0, 1, 0],
                       'v2_2':[0, 0, 0, 0, 0],
                       'v2_3':[0, 0, 0, 0, 1],
                       'v2_97':[0, 0, 0, 1, 0],
                       'v3':['good', '', '', '', 'great'],
                       'v4':['good', 'bad', 'ugly&', '', 'great*']})
    unreason(df,
             'v1', [1, 2, 3, 4],des='v1不合理值')

    multi_1(df,
            ['v2_1', 'v2_2', 'v2_3'],
            des='複選題至少選1個')

    multi_nr(df,
             ['v2_1', 'v2_2', 'v2_3'],
             ['v2_97'],
             des='複選題無反應不可與其他一起選')

    oe_list(df, 'v3', des='開放題確認')

    oe_cintent(df, 'v4', des='開放題不正常符號')
    print(result)





            


