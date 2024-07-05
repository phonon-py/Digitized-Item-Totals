import configparser
import os

import pandas as pd

from determine_business import determine_business
from msteams import MsTeams

config = configparser.ConfigParser()
config.read(r"C:\Users\041674\VSCode\デジタル化効果集計\config.ini", encoding="utf-8")

# Microsoft TeamsのWebhook URL
webhook_url = config["Teams"]["webhook_url"]

# MsTeamsクラスのインスタンスを作成
teams = MsTeams(webhook_url)

# ファイルを検索してデータフレームに追加
all_data_frames = []
for dirpath, dirnames, filenames in os.walk(config["File"]["file_path"]):
    for filename in filenames:
        # ファイル名に'水平展開'が含まれるかチェック
        if "水平展開" in filename and filename.endswith(".xlsx"):
            file_path = os.path.join(dirpath, filename)
            try:
                df = pd.read_excel(file_path, usecols=[2, 3, 4, 5], skiprows=[0, 1])
                # 欠損ある行の削除
                drop_df = df.dropna()
                # ファイル名カラムを追加
                drop_df["ファイル名"] = filename
                all_data_frames.append(drop_df)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")

# データフレームを行方向に結合
combined_df = pd.concat(all_data_frames, ignore_index=True)

# determine_business関数の実行
combined_df["事業区分"] = combined_df["導入部門"].apply(determine_business)

# pickleで保存
combined_df.to_pickle("combined_df.pickle")
# Excelで保存
combined_df.to_excel(config["User"]["save_file"])

# メンションを含む複数行テキストメッセージを送信
multi_line_txt = f"""
デジタル化水平展開シートの集計が完了しました。

以下のディレクトリを参照してください。\n
\n
{config["User"]["save_file"]}
"""

teams.create_mention_payload(
    multi_line_txt, config["User"]["id"], config["User"]["name"]
)

teams.send()
