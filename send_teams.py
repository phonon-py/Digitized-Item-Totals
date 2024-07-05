import configparser

from msteams import MsTeams

config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")

# Microsoft TeamsのWebhook URL
webhook_url = config["Teams"]["webhook_url"]

# MsTeamsクラスのインスタンスを作成
teams = MsTeams(webhook_url)

# 通常のメッセージを送信
teams.create_mention_payload("これは一般的なメッセージです。")
teams.send()

# メンションを含む複数行テキストメッセージを送信
multi_line_txt = f"""
以下のディレクトリを参照してください

{config["User"]["save_dir"]}
"""
teams.create_mention_payload(
    multi_line_txt, config["User"]["id"], config["User"]["name"]
)
teams.send()
