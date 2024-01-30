import pandas as pd
import csv
import os
from config import FILEPATH

filepath = FILEPATH

# ファイルを検索してデータフレームに追加
all_data_frames = []
for dirpath, dirnames, filenames in os.walk(filepath):
    # print(f"dirpath:{dirpath},dirnames:{dirnames}")
    for filename in filenames:
        # ファイル名に'水平展開'が含まれるかチェック
        if '水平展開' in filename and filename.endswith('.xlsx'):
            file_path = os.path.join(dirpath, filename)
            print(file_path)