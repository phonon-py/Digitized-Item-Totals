import pandas as pd
import os
from config import FILEPATH

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

# '事業区分'カラムを追加し、'導入部門'列の値に基づいて'A事業'または'B事業'と判定する関数を定義する
def determine_business(x):
    if x in ['内部統制推進室', ]:
        return '内部統制推進室'
    elif x in ['CIMS推進室',]:
        return 'CIMS推進室'
    elif x in ['管理課','人事課','施設保全課','環境管理課','安全衛生課']:
        return '人事施設部'
    elif x in ['経理課', 'IT推進課', '総合企画課']:
        return '経営管理部'
    elif x in ['品質保証第一課','品質保証第二課','品質保証第三課', '製品評価課']:
        return '品質保証部'
    elif x in ['調達第一課','調達第二課','部品QA課']:
        return '調達部'
    elif x in ['第一技術開発室','第一技術開発室', '生産技術課', '工機課']:
        return '技術開発部'
    elif x in ['生産企画部', '生産管理第一課','生産管理第二課', '製品管理課']:
        return '生産管理第一部'
    elif x in ['生産管理第三課','生産管理第四課']:
        return '生産管理第二部'
    elif x in ['CRG技術第一課','CRG技術第二課','CRG技術第三課']:
        return 'CRG技術部'
    elif x in ['CRG製造第一課', 'CRG製造第二課', 'CRG製造第三課', 'CRG製造第四課']:
        return 'CRG製造部'
    elif x in ['部品製造第一課', '部品製造第二課', '部品製造第三課', '部品製造第四課']:
        return '部品製造部'
    elif x in ['ICB製造第一課', 'ICB製造第二課']:
        return 'ICB製造部'
    elif x in ['ICB技術課','ICB組立推進課']:
        return 'ICB技術課'
    elif x in ['SEN技術第一課','SEN技術第二課','SEN技術第三課']:
        return 'SEN技術部'
    elif x in ['SEN製造第一課','SEN製造第二課','SEN製造第三課']:
        return 'SEN製造部'
    elif x in ['医療システム技術QA課', '医療システム製造課']:
        return '医療システム製造部'
    else:
        return '不明'
    
combined_df['事業区分'] = combined_df['導入部門'].apply(determine_business)


combined_df.to_pickle('combined_df.pickle')