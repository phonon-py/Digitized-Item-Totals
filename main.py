import pandas as pd
import os

filepath = 'Excelファイルパス'

# ファイルを検索してデータフレームに追加
all_data_frames = []
for dirpath, dirnames, filenames in os.walk(filepath):
    for filename in filenames:
        # ファイル名に'特定のワード'が含まれるかチェック
        if '特定のワード' in filename and filename.endswith('.xlsx'):
            file_path = os.path.join(dirpath, filename)
            try:
                # Excelファイルの特性に合わせてusecols,skiprowsを設定
                df = pd.read_excel(file_path, usecols=[2, 3, 4, 5], skiprows=[0, 1])
                
                # 欠損のある行の削除
                drop_df = df.dropna()
                
                all_data_frames.append(clean_df)                
            except Exception as e:
                print(f"Error reading {file_path}: {e}")

# データフレームを行方向に結合、ignore_index=Falseでインデックスを追加しない。
combined_df = pd.concat(all_data_frames, ignore_index=True)

# '事業区分'カラムを追加し、'a'列の値に基づいて'A事業'または'B事業'と判定する関数を定義する
def determine_business(x):
    if x in ['a', 'b']:
        return 'A事業'
    elif x in ['c', 'd']:
        return 'B事業'
    else:
        return '不明'

# apply関数を使って、'a'列の各値に対して判定関数を適用する
df['事業区分'] = df['a'].apply(determine_business)

# データ出力
export_filepath = "出力したい場所のファイルパス"
df.to_csv(export_filepath,index=False)
