import pandas as pd
import os
from config import FILEPATH
from determine_business import determine_business

filepath = FILEPATH

# ファイルを検索してデータフレームに追加
all_data_frames = []
for dirpath, dirnames, filenames in os.walk(filepath):
    for filename in filenames:
        # ファイル名に'水平展開'が含まれるかチェック
        if '水平展開' in filename and filename.endswith('.xlsx'):
            file_path = os.path.join(dirpath, filename)
            try:
                df = pd.read_excel(file_path, usecols=[2, 3, 4, 5], skiprows=[0, 1])
                # 欠損ある行の削除
                drop_df = df.dropna()
                # datetimeに変換
                drop_df['導入日'] = pd.to_datetime(drop_df['導入日'])
                clean_df = drop_df[drop_df['導入日'] >= '2024-01-01']
                clean_df['ファイル名'] = filename
                all_data_frames.append(clean_df)                
            except Exception as e:
                print(f"Error reading {file_path}: {e}")

# データフレームを行方向に結合
combined_df = pd.concat(all_data_frames, ignore_index=True)

# determine_business関数の実行
combined_df['事業区分'] = combined_df['導入部門'].apply(determine_business)

# pickleで保存
combined_df.to_pickle('combined_df.pickle')
# Excelで保存
combined_df.to_excel('combined_df.xlsx')

print('Fin')