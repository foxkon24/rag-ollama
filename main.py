# main.py - メインアプリケーションファイル（OneDrive検索機能改善版）
import os
import sys
from logger import setup_logger
from config import load_config
from flask import Flask
from teams_webhook import TeamsWebhook
from routes import register_routes
from onedrive_search import OneDriveSearch  # OneDrive検索機能

# カレントディレクトリをスクリプトのディレクトリに設定
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# ロガー設定をインポート
logger = setup_logger(script_dir)

# 設定をインポート
config = load_config(script_dir)

# Flaskアプリケーションをインポート
app = Flask(__name__)

# Teams Webhookの初期化
try:
    teams_webhook = None

    if config['TEAMS_WORKFLOW_URL']:
        teams_webhook = TeamsWebhook(config['TEAMS_WORKFLOW_URL'])
        logger.info(f"Teams Webhook初期化完了: {config['TEAMS_WORKFLOW_URL']}")
    else:
        logger.warning("TEAMS_WORKFLOW_URLが設定されていません。Teams通知機能は無効です。")

except ImportError:
    logger.error("teams_webhook モジュールが見つかりません。")
    teams_webhook = None

# OneDrive検索機能の初期化
onedrive_search = None
if config['ONEDRIVE_SEARCH_ENABLED']:
    try:
        # .env設定に基づいた初期化
        base_directory = config['ONEDRIVE_SEARCH_DIR'] if config['ONEDRIVE_SEARCH_DIR'] else None
        file_types = config['ONEDRIVE_FILE_TYPES']
        max_files = config['ONEDRIVE_MAX_FILES']

        if base_directory:
            # ディレクトリの存在確認を追加
            if os.path.exists(base_directory):
                logger.info(f"OneDrive検索ディレクトリを確認: {base_directory} (存在します)")
            else:
                logger.warning(f"OneDrive検索ディレクトリが見つかりません: {base_directory}")
                # ユーザー名を取得
                username = os.getenv("USERNAME", "owner")
                
                # 代替パスの検索を試みる - 複数のパターンを試す
                alt_paths = [
                    # 標準的なOneDriveパス
                    f"C:\\Users\\{username}\\OneDrive\\日報",
                    f"C:\\Users\\{username}\\OneDrive\\report",
                    # 企業向けOneDriveパス（共立電機製作所）
                    f"C:\\Users\\{username}\\OneDrive - 株式会社　共立電機製作所\\日報",
                    f"C:\\Users\\{username}\\OneDrive - 株式会社　共立電機製作所\\共立電機\\019.総務部\\日報"
                ]
                
                for alt_path in alt_paths:
                    if os.path.exists(alt_path):
                        logger.info(f"代替パスを使用します: {alt_path}")
                        base_directory = alt_path
                        break
                    
                # それでも見つからない場合
                if not os.path.exists(base_directory):
                    logger.warning(f"有効な代替パスが見つかりませんでした。OneDriveルートディレクトリを使用します。")
                    # OneDriveのデフォルトルート
                    onedrive_root = os.path.expanduser("~/OneDrive")
                    if os.path.exists(onedrive_root):
                        base_directory = onedrive_root
                        logger.info(f"OneDriveルートディレクトリを使用: {onedrive_root}")
                    else:
                        # ユーザーホームディレクトリ
                        home_dir = os.path.expanduser("~")
                        logger.info(f"ユーザーホームディレクトリを使用: {home_dir}")
                        base_directory = home_dir

        # OneDriveSearch インスタンスを初期化
        onedrive_search = OneDriveSearch(
            base_directory=base_directory,
            file_types=file_types,
            max_results=max_files
        )
        logger.info(f"OneDrive検索機能を初期化しました: {base_directory}")

    except ImportError:
        logger.error("onedrive_search モジュールが見つかりません。OneDrive検索機能は無効です。")

    except Exception as e:
        logger.error(f"OneDrive検索機能の初期化中にエラーが発生しました: {str(e)}")

else:
    logger.info("OneDrive検索機能は.envの設定により無効化されています")

# ルートの登録（OneDrive検索機能を渡す）
register_routes(app, config, teams_webhook, onedrive_search)

# アプリケーション実行
if __name__ == '__main__':
    try:
        logger.info("Ollama連携アプリケーションを起動します")
        logger.info(f"環境: {os.getenv('FLASK_ENV', 'production')}")
        logger.info(f"デバッグモード: {config['DEBUG']}")
        logger.info(f"Ollama URL: {config['OLLAMA_URL']}")
        logger.info(f"Ollama モデル: {config['OLLAMA_MODEL']}")
        logger.info(f"Teams Webhook URL: {config['TEAMS_WORKFLOW_URL'] if config['TEAMS_WORKFLOW_URL'] else '未設定'}")
        
        if onedrive_search:
            search_path = onedrive_search.base_directory
            logger.info(f"OneDrive検索: 有効 (検索パス: {search_path})")
        else:
            logger.info("OneDrive検索: 無効")

        # アプリケーションを起動
        logger.info(f"アプリケーションを起動: {config['HOST']}:{config['PORT']}")
        
        # デバッガー環境での実行に対応
        try:
            # Flask 2.2.0以降の互換性対応
            app.run(
                host=config['HOST'], 
                port=config['PORT'], 
                debug=config['DEBUG'],
                use_reloader=False  # デバッガーと競合する可能性があるため無効化
            )
        except TypeError:
            # 古いバージョンのFlaskに対する互換性対応
            app.run(
                host=config['HOST'], 
                port=config['PORT'], 
                debug=config['DEBUG']
            )

    except Exception as e:
        logger.error(f"アプリケーション起動中にエラーが発生しました: {str(e)}")
        logger.error("詳細なエラー情報:", exc_info=True)  # 詳細なスタックトレースをログに記録
