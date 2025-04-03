# main.py - メインアプリケーションファイル（OneDrive検索機能追加版）
import os
import sys
from logger import setup_logger
from config import load_config
from flask import Flask
from teams_webhook import TeamsWebhook
from routes import register_routes
from onedrive_search import OneDriveSearch  # 新たに追加したOneDrive検索機能

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

        onedrive_search = OneDriveSearch(
            base_directory=base_directory,
            file_types=file_types,
            max_results=max_files
        )
        logger.info("OneDrive検索機能を初期化しました")

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
        logger.info(f"OneDrive検索: {'有効' if onedrive_search else '無効'}")

        # アプリケーションを起動
        logger.info(f"アプリケーションを起動: {config['HOST']}:{config['PORT']}")
        app.run(host=config['HOST'], port=config['PORT'], debug=config['DEBUG'])

    except Exception as e:
        logger.error(f"アプリケーション起動中にエラーが発生しました: {str(e)}")
        raise
