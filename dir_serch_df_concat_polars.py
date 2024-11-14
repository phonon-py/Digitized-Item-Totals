import configparser
import os

import polars as pl

from determine_business import determine_business
from msteams import MsTeams

config = configparser.ConfigParser()
config.read(r"C:\\Users\\041674\\VSCode\\デジタル化効果集計\\config.ini", encoding="utf-8")

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
                df = pl.read_excel(file_path, sheet_name=0, skip_rows=2, columns=['C', 'D', 'E', 'F'])
                # 欠損ある行の削除
                df = df.drop_nulls()
                # ファイル名カラムを追加
                df = df.with_column(pl.lit(filename).alias("ファイル名"))
                all_data_frames.append(df)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")

# データフレームを行方向に結合
combined_df = pl.concat(all_data_frames)

# determine_business関数の実行
combined_df = combined_df.with_column(
    combined_df["導入部門"].apply(determine_business).alias("事業区分")
)

# pickleで保存
combined_df.write_pickle("combined_df.pickle")
# Excelで保存
combined_df.write_excel(config["User"]["save_file"])

# メンションを含む複数行テキストメッセージを送信
multi_line_txt = f"""
デジタル化水平展開シートの集計が完了しました。

以下のディレクトリを参照してください。\n\n{config["User"]["save_file"]}
"""

teams.create_mention_payload(
    multi_line_txt, config["User"]["id"], config["User"]["name"]
)

teams.send()
