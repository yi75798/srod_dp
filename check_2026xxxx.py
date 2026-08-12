#!/usr/bin/python
# -*- encoding: utf-8 -*-
# File    :   check_2026xxxx.py
# Time    :   2026/08/12 11:18:52
# Author  :   Hsu, Liang-Yi 
# Email:   yi75798@gmail.com
# Description :

import pandas as pd
import numpy as np
from process.data_check import *

### 定義路徑
## 原始資料檔
raw_data_path = r'C:\Users\yi5011\Documents\GitHub\srod_dp\data_s_test.sav'
## 輸出檔路徑
output_path = r'C:\Users\yi5011\Documents\GitHub\srod_dp\output0812.xlsx'

### 讀取檔案
df = data_import(raw_data_path)

### 定義數值遺漏值
df = def_na(df)

#### 建立固定變項
Id = 'id' # 樣本編號
wno = 'wno' # 會員編號
gender = 'gender2026' # 性別
birth = 'birth2026' # 出生年

#### 建立輸出檔
## 不符合品
result = pd.DataFrame(columns=['檢核日期', '樣本編號', 'wno', '性別', '年次', '變項名稱', '不符合說明'])
## 開放題清單
oe_answer = pd.DataFrame(columns=['檢核日期', '樣本編號', 'wno', '性別', '年次', '變項名稱', '不符合說明'])

### 基本檢核開始
## 連續不合理
# vA1
unreason_con(df, result, 'vA1', [0, 9999999], [])

## 跳答
# vB1-vB1U
skip(df, result, 'vB1', [1], 'vB1U', [9999999996])
# vB1-vB1D
skip(df, result, 'vB1', [2], 'vB1D', [9999999996])

## 複選數量檢核
# vF1
F1 = ['vF1m'+str(n) for n in range(1, 12)]
multi_1(df, result, F1)

## 開放題列出
# vC3o7
oe_list(df, oe_answer, 'vC3o7')

## 開放題不正常符號列出
# vC3o7
oe_abnormal(df, oe_answer, 'vC3o7')

#----------------------------------------------------------------------
#### 特殊檢核開始








#----------------------------------------------------------------------
### 輸出報表
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    result.to_excel(writer, sheet_name='不符合品報表', index=0)
    oe_answer.to_excel(writer, sheet_name='開放題答案確認', index=0)