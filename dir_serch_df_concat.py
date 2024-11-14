import configparser
import logging
import os
from logging.handlers import RotatingFileHandler
import unicodedata  # 文字列の正規化に使用

import pandas as pd

from determine_business import determine_business

from sqlalchemy import create_engine, text
from logging.handlers import RotatingFileHandler

def setup_logging(log_file='script.log', max_size=5*1024*1024, backup_count=3):
    """ログの設定を行う関数"""
    # ログフォーマットの設定
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # ルートロガーの設定
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # ファイルハンドラの設定（ローテーション付き）
    file_handler = RotatingFileHandler(
        log_file, maxBytes=max_size, backupCount=backup_count
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(log_format))

    # コンソールハンドラの設定
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter(log_format))

    # ハンドラをルートロガーに追加
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

def get_organization_data(engine):
    """Oracleから組織データを取得して結合する関数"""
    logger = logging.getLogger(__name__)
    
    try:        
        with engine.connect() as conn:
            # 各テーブルから最新データを取得
            df_org = pd.read_sql_query(text("SELECT * FROM WPUSER.PA_M_SOSIKI_NAME"), conn)
            df_business_center = pd.read_sql_query(
                text(f"""
                SELECT * FROM WPUSER.PA_M_PROD_DIV_DEPT 
                """), conn)
            df_business_name = pd.read_sql_query(
                text(f"""
                SELECT * FROM WPUSER.PA_M_PROD_DIV 
                """), conn)

        # データ型の統一
        df_org['cd_sosiki'] = df_org['cd_sosiki'].astype(str)
        df_business_center['cd_dept'] = df_business_center['cd_dept'].astype(str)
        df_business_center['cd_div'] = df_business_center['cd_div'].astype(str)

        # データ結合
        df_merged = pd.merge(
            df_business_center, 
            df_business_name[['cd_prod_div', 'nm_prod_div_n']], 
            on='cd_prod_div', 
            how='left'
        )

        df_merged = pd.merge(
            df_merged,
            df_org[['cd_sosiki', 'nm_sosiki']],  # 必要な列のみを選択
            left_on='cd_dept',
            right_on='cd_sosiki',
            how='left'
        )

        df_merged = pd.merge(
            df_merged,
            df_org[['cd_sosiki', 'nm_sosiki']].rename(columns={'cd_sosiki': 'cd_div', 'nm_sosiki': 'nm_div'}),
            on='cd_div',
            how='left'
        )

        # 必要な列の選択と名前変更
        result = df_merged[['cd_sosiki', 'nm_prod_div_n', 'nm_div', 'nm_sosiki']].rename(columns={
            'nm_prod_div_n': '事業部門',
            'nm_div': '部名',
            'nm_sosiki': '課名',
            'cd_sosiki': 'cd_sosiki'  # コードも保持
        })

        # 重複削除とソート
        result = result.drop_duplicates().sort_values(['事業部門', '部名', '課名'])
        result = result.fillna('未設定')
        
        logger.info("組織データの取得と結合が完了しました")
        return result
        
    except Exception as e:
        logger.exception("組織データの取得中にエラーが発生しました")
        raise

# 英字を半角に変換する関数
def convert_alpha_to_half_width(text):
    if pd.isna(text):
        return text
    # 文字列に変換
    text = str(text)
    # 英字を半角に変換
    return unicodedata.normalize('NFKC', text)

def main():
    logger = logging.getLogger(__name__)
    logger.info("スクリプトを開始します")
    

    try:
        # 設定ファイルの読み込み（.env.ini と config.ini を統合したファイル）
        config = configparser.ConfigParser()
        # configファイルの読み込み確認用にデバッグログを追加
        config_path = r"C:\Users\041674\Python3.12.6\デジタル化効果集計\config.ini"
        if config.read(config_path, encoding="utf-8"):
            logger.info(f"設定ファイルを読み込みました: {config.sections()}")
        else:
            logger.error(f"設定ファイル {config_path} を読み込めませんでした")
            raise FileNotFoundError(f"設定ファイル {config_path} が見つかりません")

        # ORACLE_ACH_CSYDBセクションの存在確認
        if "ORACLE_ACH_CSYDB" not in config:
            logger.error("ORACLE_ACH_CSYDBセクションが設定ファイルにありません")
            raise KeyError("ORACLE_ACH_CSYDBセクションが見つかりません")

        # Oracle接続設定（例外処理を追加）
        try:
            username = config["ORACLE_ACH_CSYDB"]["DB_USERNAME"]
            password = config["ORACLE_ACH_CSYDB"]["DB_PASSWORD"]
            hostname = config["ORACLE_ACH_CSYDB"]["DB_HOSTNAME"]
            port = config["ORACLE_ACH_CSYDB"]["DB_PORT"]
            service_name = config["ORACLE_ACH_CSYDB"]["DB_SERVICE_NAME"]
        except KeyError as e:
            logger.error(f"必要な設定項目が見つかりません: {e}")
            raise
        
        # Oracle接続文字列
        DATABASE_URL = f"oracle+cx_oracle://{username}:{password}@{hostname}:{port}/?service_name={service_name}"
        
        # エンジン作成
        engine = create_engine(DATABASE_URL, echo=False)
        
        # 組織データの取得
        org_data = get_organization_data(engine)
        logger.info(f"Oracleから最新のデータを取得することができました:\n{print(org_data)}")

        all_data_frames = []
        file_count = 0
        processed_count = 0

        for dirpath, dirnames, filenames in os.walk(config["File"]["file_path"]):
            for filename in filenames:
                file_count += 1
                if "水平展開" in filename and filename.endswith(".xlsx"):
                    file_path = os.path.join(dirpath, filename)
                    logger.info(f"処理中のファイル: {file_path}")
                    try:
                        df = pd.read_excel(file_path, usecols=[2, 3, 4, 5], skiprows=[0, 1])
                        logger.debug(f"読み込んだデータフレームの形状: {df.shape}")

                        drop_df = df.dropna()
                        logger.debug(f"欠損値削除後のデータフレームの形状: {drop_df.shape}")

                        drop_df["ファイル名"] = filename
                        all_data_frames.append(drop_df)
                        processed_count += 1
                        logger.info(f"ファイルの処理が成功しました: {filename}")
                    except Exception as e:
                        logger.exception(f"ファイル {file_path} の読み込み中にエラーが発生しました")

        logger.info(f"合計 {file_count} ファイルをスキャンし、{processed_count} ファイルを処理しました")

        if not all_data_frames:
            logger.warning("処理されたデータフレームがありません")
            return

        combined_df = pd.concat(all_data_frames, ignore_index=True)
        logger.info(f"結合されたデータフレームの形状: {combined_df.shape}")

        # 両方のデータフレームで英字を半角に変換
        combined_df['導入部門'] = combined_df['導入部門'].apply(convert_alpha_to_half_width)
        org_data['課名'] = org_data['課名'].apply(convert_alpha_to_half_width)

        # TODO git pushして前のバージョンに戻す
        combined_df = pd.merge(
            combined_df,
            org_data[['事業部門', '部名', '課名']],
            left_on='導入部門',
            right_on='課名',
            how='left'
        )

        # マッチングできなかった行に対する処理
        combined_df.loc[combined_df['事業部門'].isna(), '事業部門'] = '不明'
        combined_df.loc[combined_df['部名'].isna(), '部名'] = '不明'
        logger.info("事業区分を決定しました")

        save_path = config["User"]["save_file"]
        combined_df.to_excel(save_path, index=False)
        logger.info(f"結果をExcelファイルに保存しました(インデックスなし): {save_path}")

    except Exception as e:
        logger.exception("スクリプト実行中に予期せぬエラーが発生しました")
    else:
        logger.info("スクリプトが正常に完了しました")

if __name__ == "__main__":
    setup_logging()
    main()