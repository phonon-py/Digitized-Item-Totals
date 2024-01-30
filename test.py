import pandas as pd
import csv
import os
from config import FILEPATH

filepath = r'\\ach-lcfile01\Share001\4940_デジタル化案件管理\T22-001\【T22-001】 水平展開効果算出シート.xlsx'

df = pd.read_excel(filepath,usecols=[1,2,3,4,5],skiprows=[0,1])

# 欠損ある行の削除
drop_df = df.dropna()
# datetimeに変換
drop_df['導入日'] = pd.to_datetime(drop_df['導入日'])

try:
    clean_df = drop_df[drop_df['導入日'] >= '2023-01-01']
    clean_df['filepath'] = filepath
    print(clean_df)
except:
    print('データがありません')
    pass
